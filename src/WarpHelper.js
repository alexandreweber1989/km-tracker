// ═══════════════════════════════════════════════════════════════════════════
// WARP HELPER (OpenCV Edge Detection + Perspective Transform)
// Scanner profissional com detecção de bordas via Visão Computacional C++
// ═══════════════════════════════════════════════════════════════════════════

export async function autoCropDocument(imageElement) {
  // Fallback se o OpenCV não foi carregado ainda pela internet
  if (typeof cv === 'undefined' || !cv.Mat) {
    console.warn("OpenCV não carregado a tempo, usando fallback normal.");
    return await fallbackWarp(imageElement);
  }

  try {
    const sw = imageElement.naturalWidth || imageElement.width;
    const sh = imageElement.naturalHeight || imageElement.height;

    // Redimensionar apenas para o algoritmo de detecção ser ultra-rápido (< 50ms)
    const MAX_DIM = 800;
    let scale = 1;
    let procWidth = sw;
    let procHeight = sh;

    if (sw > MAX_DIM || sh > MAX_DIM) {
      if (sw > sh) {
        scale = MAX_DIM / sw;
        procWidth = MAX_DIM;
        procHeight = sh * scale;
      } else {
        scale = MAX_DIM / sh;
        procHeight = MAX_DIM;
        procWidth = sw * scale;
      }
    }

    const canvas = document.createElement('canvas');
    canvas.width = procWidth;
    canvas.height = procHeight;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(imageElement, 0, 0, procWidth, procHeight);

    let src = cv.imread(canvas);
    let gray = new cv.Mat();
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY, 0);
    
    // Suavização para matar ruído (textura da mesa, marcas d'água)
    let blur = new cv.Mat();
    cv.GaussianBlur(gray, blur, new cv.Size(5, 5), 0, 0, cv.BORDER_DEFAULT);

    // Canny edge detection (encontra as linhas)
    let edges = new cv.Mat();
    cv.Canny(blur, edges, 75, 200, 3, false);

    // Encontrar os contornos das linhas fechadas
    let contours = new cv.MatVector();
    let hierarchy = new cv.Mat();
    cv.findContours(edges, contours, hierarchy, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE);

    let maxArea = 0;
    let bestApprox = null;

    // Procurar o maior contorno com exatos 4 lados (o papel do recibo)
    for (let i = 0; i < contours.size(); ++i) {
      let cnt = contours.get(i);
      let area = cv.contourArea(cnt);
      
      // O documento deve preencher pelo menos 10% da foto
      if (area > (procWidth * procHeight) * 0.1) {
        let peri = cv.arcLength(cnt, true);
        let approx = new cv.Mat();
        // Opsilon de 0.02 permite que o papel não seja 100% liso nas bordas
        cv.approxPolyDP(cnt, approx, 0.02 * peri, true);
        
        if (approx.rows === 4 && area > maxArea) {
          maxArea = area;
          if (bestApprox) bestApprox.delete();
          bestApprox = approx.clone();
        }
        approx.delete();
      }
      cnt.delete();
    }

    if (bestApprox) {
      // Mapear pontos devolta para a resolução original (HD)
      let points = [];
      for (let i = 0; i < 4; i++) {
        points.push({
          x: bestApprox.data32S[i * 2] / scale,
          y: bestApprox.data32S[i * 2 + 1] / scale
        });
      }
      points = orderPoints(points);

      // Limpar memória WebAssembly
      src.delete(); gray.delete(); blur.delete(); edges.delete(); contours.delete(); hierarchy.delete(); bestApprox.delete();

      console.log("Recibo localizado via OpenCV!");
      return await warpPerspectiveCV(imageElement, points);
    } else {
      console.warn("Nenhum papel com 4 pontas foi detectado claramente pelo OpenCV. Usando fallback.");
      src.delete(); gray.delete(); blur.delete(); edges.delete(); contours.delete(); hierarchy.delete();
      return await fallbackWarp(imageElement);
    }
  } catch(e) {
    console.error("OpenCV AutoCrop Error:", e);
    return await fallbackWarp(imageElement);
  }
}

// Order: Top-Left, Top-Right, Bottom-Right, Bottom-Left
function orderPoints(pts) {
  pts.sort((a,b) => (a.x + a.y) - (b.x + b.y));
  const tl = pts[0];
  const br = pts[3];
  
  const remain = [pts[1], pts[2]];
  remain.sort((a,b) => (a.y - a.x) - (b.y - b.x));
  const tr = remain[0]; // Top-right will have smaller y, larger x -> smaller (y-x)
  const bl = remain[1]; // Bottom-left will have larger y, smaller x -> larger (y-x)
  
  return [tl, tr, br, bl];
}

async function warpPerspectiveCV(imageElement, srcCorners) {
  const sw = imageElement.naturalWidth || imageElement.width;
  const sh = imageElement.naturalHeight || imageElement.height;

  const topWidth = Math.hypot(srcCorners[1].x - srcCorners[0].x, srcCorners[1].y - srcCorners[0].y);
  const bottomWidth = Math.hypot(srcCorners[2].x - srcCorners[3].x, srcCorners[2].y - srcCorners[3].y);
  const leftHeight = Math.hypot(srcCorners[3].x - srcCorners[0].x, srcCorners[3].y - srcCorners[0].y);
  const rightHeight = Math.hypot(srcCorners[2].x - srcCorners[1].x, srcCorners[2].y - srcCorners[1].y);
  
  let outWidth = Math.max(topWidth, bottomWidth);
  let outHeight = Math.max(leftHeight, rightHeight);

  // Limite da imagem final
  const MAX_DIMENSION = 1600;
  if (outWidth > outHeight && outWidth > MAX_DIMENSION) {
    outHeight = outHeight * (MAX_DIMENSION / outWidth);
    outWidth = MAX_DIMENSION;
  } else if (outHeight > MAX_DIMENSION) {
    outWidth = outWidth * (MAX_DIMENSION / outHeight);
    outHeight = MAX_DIMENSION;
  }
  
  outWidth = Math.floor(outWidth);
  outHeight = Math.floor(outHeight);

  const canvas = document.createElement('canvas');
  canvas.width = sw;
  canvas.height = sh;
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.drawImage(imageElement, 0, 0, sw, sh);

  let src = cv.imread(canvas);
  let dst = new cv.Mat();
  let dsize = new cv.Size(outWidth, outHeight);
  
  let srcTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
    srcCorners[0].x, srcCorners[0].y,
    srcCorners[1].x, srcCorners[1].y,
    srcCorners[2].x, srcCorners[2].y,
    srcCorners[3].x, srcCorners[3].y
  ]);
  
  let dstTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
    0, 0,
    outWidth, 0,
    outWidth, outHeight,
    0, outHeight
  ]);
  
  let M = cv.getPerspectiveTransform(srcTri, dstTri);
  cv.warpPerspective(src, dst, M, dsize, cv.INTER_LINEAR, cv.BORDER_CONSTANT, new cv.Scalar());
  
  const finalCanvas = document.createElement('canvas');
  finalCanvas.width = outWidth;
  finalCanvas.height = outHeight;
  cv.imshow(finalCanvas, dst);
  
  src.delete(); dst.delete(); M.delete(); srcTri.delete(); dstTri.delete();

  return new Promise(resolve => {
    finalCanvas.toBlob(blob => resolve({ blob, dataUrl: finalCanvas.toDataURL('image/jpeg', 0.9) }), 'image/jpeg', 0.9);
  });
}

async function fallbackWarp(imageElement) {
  const sw = imageElement.naturalWidth || imageElement.width;
  const sh = imageElement.naturalHeight || imageElement.height;
  
  let outW = sw, outH = sh;
  const MAX = 1600;
  if (outW > MAX || outH > MAX) {
    if (outW > outH) { outH = outH * (MAX / outW); outW = MAX; }
    else { outW = outW * (MAX / outH); outH = MAX; }
  }
  
  outW = Math.floor(outW);
  outH = Math.floor(outH);
  
  const canvas = document.createElement('canvas');
  canvas.width = outW;
  canvas.height = outH;
  canvas.getContext('2d').drawImage(imageElement, 0, 0, outW, outH);
  
  return new Promise(resolve => {
    canvas.toBlob(blob => resolve({ blob, dataUrl: canvas.toDataURL('image/jpeg', 0.9) }), 'image/jpeg', 0.9);
  });
}
