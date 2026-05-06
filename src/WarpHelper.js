// ═══════════════════════════════════════════════════════════════════════════
// WARP HELPER (Perspective Transform / Deskew)
// Implementação nativa de Homografia para alinhar e recortar imagens
// ═══════════════════════════════════════════════════════════════════════════

function solveLinearSystem(A, b) {
  const n = A.length;
  for (let i = 0; i < n; i++) {
    let maxRow = i;
    for (let k = i + 1; k < n; k++) {
      if (Math.abs(A[k][i]) > Math.abs(A[maxRow][i])) maxRow = k;
    }
    [A[i], A[maxRow]] = [A[maxRow], A[i]];
    [b[i], b[maxRow]] = [b[maxRow], b[i]];
    
    const p = A[i][i];
    if (Math.abs(p) < 1e-10) return null;
    for (let j = i; j < n; j++) A[i][j] /= p;
    b[i] /= p;
    
    for (let k = 0; k < n; k++) {
      if (k === i) continue;
      const factor = A[k][i];
      for (let j = i; j < n; j++) A[k][j] -= factor * A[i][j];
      b[k] -= factor * b[i];
    }
  }
  return b;
}

function getHomographyMatrix(src, dst) {
  const A = [];
  const b = [];
  for (let i = 0; i < 4; i++) {
    const sx = src[i].x;
    const sy = src[i].y;
    const dx = dst[i].x;
    const dy = dst[i].y;
    A.push([sx, sy, 1, 0, 0, 0, -dx * sx, -dx * sy]);
    b.push(dx);
    A.push([0, 0, 0, sx, sy, 1, -dy * sx, -dy * sy]);
    b.push(dy);
  }
  return solveLinearSystem(A, b);
}

function applyHomography(H, x, y) {
  const denominator = H[6] * x + H[7] * y + 1;
  return {
    x: (H[0] * x + H[1] * y + H[2]) / denominator,
    y: (H[3] * x + H[4] * y + H[5]) / denominator
  };
}

/**
 * Aplica o Perspective Transform na imagem.
 * @param {HTMLImageElement} imageElement Imagem original carregada
 * @param {Array} srcCorners Array de 4 coordenadas [{x,y}, ...] em pixels da imagem original (ordem: Top-Left, Top-Right, Bottom-Right, Bottom-Left)
 * @returns {Promise<{blob: Blob, dataUrl: string}>}
 */
export async function warpPerspective(imageElement, srcCorners) {
  const sw = imageElement.naturalWidth || imageElement.width;
  const sh = imageElement.naturalHeight || imageElement.height;

  const topWidth = Math.hypot(srcCorners[1].x - srcCorners[0].x, srcCorners[1].y - srcCorners[0].y);
  const bottomWidth = Math.hypot(srcCorners[2].x - srcCorners[3].x, srcCorners[2].y - srcCorners[3].y);
  const leftHeight = Math.hypot(srcCorners[3].x - srcCorners[0].x, srcCorners[3].y - srcCorners[0].y);
  const rightHeight = Math.hypot(srcCorners[2].x - srcCorners[1].x, srcCorners[2].y - srcCorners[1].y);
  
  let outWidth = Math.max(topWidth, bottomWidth);
  let outHeight = Math.max(leftHeight, rightHeight);

  const MAX_DIMENSION = 1200;
  if (outWidth > outHeight && outWidth > MAX_DIMENSION) {
    outHeight = outHeight * (MAX_DIMENSION / outWidth);
    outWidth = MAX_DIMENSION;
  } else if (outHeight > MAX_DIMENSION) {
    outWidth = outWidth * (MAX_DIMENSION / outHeight);
    outHeight = MAX_DIMENSION;
  }
  
  outWidth = Math.floor(outWidth);
  outHeight = Math.floor(outHeight);

  const srcCanvas = document.createElement('canvas');
  srcCanvas.width = sw;
  srcCanvas.height = sh;
  const srcCtx = srcCanvas.getContext('2d', { willReadFrequently: true });
  srcCtx.drawImage(imageElement, 0, 0, sw, sh);
  const srcData = srcCtx.getImageData(0, 0, sw, sh);

  const dstCorners = [
    {x: 0, y: 0},
    {x: outWidth, y: 0},
    {x: outWidth, y: outHeight},
    {x: 0, y: outHeight}
  ];

  const H_inv = getHomographyMatrix(dstCorners, srcCorners);
  if (!H_inv) throw new Error("A matriz de homografia falhou (pontos inválidos).");

  const dstCanvas = document.createElement('canvas');
  dstCanvas.width = outWidth;
  dstCanvas.height = outHeight;
  const dstCtx = dstCanvas.getContext('2d');
  const dstData = dstCtx.createImageData(outWidth, outHeight);

  for (let y = 0; y < outHeight; y++) {
    for (let x = 0; x < outWidth; x++) {
      const srcPt = applyHomography(H_inv, x, y);
      const sx = Math.floor(srcPt.x);
      const sy = Math.floor(srcPt.y);
      
      const dstIdx = (y * outWidth + x) * 4;
      if (sx >= 0 && sx < sw && sy >= 0 && sy < sh) {
        const srcIdx = (sy * sw + sx) * 4;
        dstData.data[dstIdx] = srcData.data[srcIdx];
        dstData.data[dstIdx+1] = srcData.data[srcIdx+1];
        dstData.data[dstIdx+2] = srcData.data[srcIdx+2];
        dstData.data[dstIdx+3] = 255;
      } else {
        dstData.data[dstIdx+3] = 0; 
      }
    }
  }

  dstCtx.putImageData(dstData, 0, 0);

  return new Promise(resolve => {
    dstCanvas.toBlob(blob => resolve({ blob, dataUrl: dstCanvas.toDataURL('image/jpeg', 0.9) }), 'image/jpeg', 0.9);
  });
}
