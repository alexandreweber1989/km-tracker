// ═══════════════════════════════════════════════════════════════════════════
// WARP HELPER (OpenCV Edge Detection + Perspective Transform + Enhancement)
// Scanner profissional com detecção de bordas + pós-processamento de imagem
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Pipeline completo:
 * 1. Detecção de bordas (Canny Edge Detection)
 * 2. Encontrar contorno do documento (4 cantos)
 * 3. Perspective Warp (deskew)
 * 4. Enhancement (CLAHE + Unsharp Mask)
 */
export async function autoCropDocument(imageElement) {
  if (typeof cv === 'undefined' || !cv.Mat) {
    console.warn("OpenCV não carregado, usando fallback.");
    return await fallbackWarp(imageElement);
  }

  try {
    const sw = imageElement.naturalWidth || imageElement.width;
    const sh = imageElement.naturalHeight || imageElement.height;

    // Redimensionar para detecção rápida (< 50ms no S25 Ultra)
    const MAX_DIM = 1024;
    let scale = 1;
    let procWidth = sw;
    let procHeight = sh;

    if (sw > MAX_DIM || sh > MAX_DIM) {
      if (sw > sh) {
        scale = MAX_DIM / sw;
        procWidth = MAX_DIM;
        procHeight = Math.round(sh * scale);
      } else {
        scale = MAX_DIM / sh;
        procHeight = MAX_DIM;
        procWidth = Math.round(sw * scale);
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

    // Suavização para matar ruído
    let blur = new cv.Mat();
    cv.GaussianBlur(gray, blur, new cv.Size(5, 5), 0, 0, cv.BORDER_DEFAULT);

    // Canny edge detection
    let edges = new cv.Mat();
    cv.Canny(blur, edges, 50, 150, 3, false);

    // Dilatar para conectar bordas quebradas
    let dilated = new cv.Mat();
    let kernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(3, 3));
    cv.dilate(edges, dilated, kernel);
    kernel.delete();

    // Encontrar contornos
    let contours = new cv.MatVector();
    let hierarchy = new cv.Mat();
    cv.findContours(dilated, contours, hierarchy, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE);

    let maxArea = 0;
    let bestApprox = null;

    for (let i = 0; i < contours.size(); ++i) {
      let cnt = contours.get(i);
      let area = cv.contourArea(cnt);

      if (area > (procWidth * procHeight) * 0.08) {
        let peri = cv.arcLength(cnt, true);
        let approx = new cv.Mat();
        cv.approxPolyDP(cnt, approx, 0.02 * peri, true);

        if (approx.rows === 4 && area > maxArea) {
          // Verificar se é convexo
          if (cv.isContourConvex(approx)) {
            maxArea = area;
            if (bestApprox) bestApprox.delete();
            bestApprox = approx.clone();
          }
        }
        approx.delete();
      }
      cnt.delete();
    }

    if (bestApprox) {
      let points = [];
      for (let i = 0; i < 4; i++) {
        points.push({
          x: bestApprox.data32S[i * 2] / scale,
          y: bestApprox.data32S[i * 2 + 1] / scale
        });
      }
      points = orderPoints(points);

      src.delete(); gray.delete(); blur.delete(); edges.delete();
      dilated.delete(); contours.delete(); hierarchy.delete(); bestApprox.delete();

      console.log("✅ Documento detectado via OpenCV!");
      const warped = await warpPerspectiveCV(imageElement, points);
      return await enhanceDocument(warped);
    } else {
      console.warn("⚠️ Contorno não encontrado, usando fallback com enhancement.");
      src.delete(); gray.delete(); blur.delete(); edges.delete();
      dilated.delete(); contours.delete(); hierarchy.delete();
      const fb = await fallbackWarp(imageElement);
      return await enhanceDocument(fb);
    }
  } catch (e) {
    console.error("OpenCV AutoCrop Error:", e);
    return await fallbackWarp(imageElement);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// PÓS-PROCESSAMENTO DE IMAGEM (CLAHE + Unsharp Mask)
// ═══════════════════════════════════════════════════════════════════════════

async function enhanceDocument(result) {
  if (typeof cv === 'undefined' || !cv.Mat) return result;

  try {
    const img = new Image();
    img.src = result.dataUrl;
    await new Promise(r => { img.onload = r; img.onerror = r; });

    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth || img.width;
    canvas.height = img.naturalHeight || img.height;
    canvas.getContext('2d', { willReadFrequently: true }).drawImage(img, 0, 0);

    let src = cv.imread(canvas);

    // 1. Converter para LAB para processar luminância separadamente
    let lab = new cv.Mat();
    cv.cvtColor(src, lab, cv.COLOR_RGBA2RGB);
    let labConverted = new cv.Mat();
    cv.cvtColor(lab, labConverted, cv.COLOR_RGB2Lab);
    lab.delete();

    // 2. Separar canais
    let channels = new cv.MatVector();
    cv.split(labConverted, channels);
    let lChannel = channels.get(0);

    // 3. CLAHE (Contrast Limited Adaptive Histogram Equalization)
    // Isso melhora dramaticamente a legibilidade do texto
    let clahe = new cv.CLAHE(3.0, new cv.Size(8, 8));
    let enhanced = new cv.Mat();
    clahe.apply(lChannel, enhanced);
    clahe.delete();

    // Substituir canal L pelo enhanced
    channels.set(0, enhanced);
    let mergedLab = new cv.Mat();
    cv.merge(channels, mergedLab);

    // Converter de volta para RGB
    let rgbEnhanced = new cv.Mat();
    cv.cvtColor(mergedLab, rgbEnhanced, cv.COLOR_Lab2RGB);

    // 4. Unsharp Mask (aumenta nitidez do texto)
    let blurred = new cv.Mat();
    cv.GaussianBlur(rgbEnhanced, blurred, new cv.Size(0, 0), 2.0);
    let sharpened = new cv.Mat();
    cv.addWeighted(rgbEnhanced, 1.5, blurred, -0.5, 0, sharpened);
    blurred.delete();

    // Renderizar resultado final
    let finalRgba = new cv.Mat();
    cv.cvtColor(sharpened, finalRgba, cv.COLOR_RGB2RGBA);

    const finalCanvas = document.createElement('canvas');
    finalCanvas.width = canvas.width;
    finalCanvas.height = canvas.height;
    cv.imshow(finalCanvas, finalRgba);

    // Limpar memória
    src.delete(); labConverted.delete(); lChannel.delete();
    enhanced.delete(); mergedLab.delete(); rgbEnhanced.delete();
    sharpened.delete(); finalRgba.delete(); channels.delete();

    console.log("✅ Enhancement aplicado (CLAHE + Unsharp Mask)");

    return new Promise(resolve => {
      finalCanvas.toBlob(blob => resolve({
        blob,
        dataUrl: finalCanvas.toDataURL('image/jpeg', 0.92)
      }), 'image/jpeg', 0.92);
    });
  } catch (e) {
    console.warn("Enhancement falhou, retornando original:", e);
    return result;
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// ORDENAÇÃO DE PONTOS
// ═══════════════════════════════════════════════════════════════════════════

function orderPoints(pts) {
  // Ordena: Top-Left, Top-Right, Bottom-Right, Bottom-Left
  const sorted = [...pts].sort((a, b) => (a.x + a.y) - (b.x + b.y));
  const tl = sorted[0];
  const br = sorted[3];
  const remain = [sorted[1], sorted[2]];
  remain.sort((a, b) => (a.y - a.x) - (b.y - b.x));
  const tr = remain[0];
  const bl = remain[1];
  return [tl, tr, br, bl];
}

// ═══════════════════════════════════════════════════════════════════════════
// PERSPECTIVE WARP (OpenCV nativo)
// ═══════════════════════════════════════════════════════════════════════════

async function warpPerspectiveCV(imageElement, srcCorners) {
  const sw = imageElement.naturalWidth || imageElement.width;
  const sh = imageElement.naturalHeight || imageElement.height;

  const topWidth = Math.hypot(srcCorners[1].x - srcCorners[0].x, srcCorners[1].y - srcCorners[0].y);
  const bottomWidth = Math.hypot(srcCorners[2].x - srcCorners[3].x, srcCorners[2].y - srcCorners[3].y);
  const leftHeight = Math.hypot(srcCorners[3].x - srcCorners[0].x, srcCorners[3].y - srcCorners[0].y);
  const rightHeight = Math.hypot(srcCorners[2].x - srcCorners[1].x, srcCorners[2].y - srcCorners[1].y);

  let outWidth = Math.max(topWidth, bottomWidth);
  let outHeight = Math.max(leftHeight, rightHeight);

  // S25 Ultra pode lidar com resolução alta
  const MAX_DIMENSION = 2400;
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
  canvas.getContext('2d', { willReadFrequently: true }).drawImage(imageElement, 0, 0, sw, sh);

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
    finalCanvas.toBlob(blob => resolve({
      blob,
      dataUrl: finalCanvas.toDataURL('image/jpeg', 0.92)
    }), 'image/jpeg', 0.92);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// FALLBACK (sem OpenCV)
// ═══════════════════════════════════════════════════════════════════════════

async function fallbackWarp(imageElement) {
  const sw = imageElement.naturalWidth || imageElement.width;
  const sh = imageElement.naturalHeight || imageElement.height;

  let outW = sw, outH = sh;
  const MAX = 2400;
  if (outW > MAX || outH > MAX) {
    if (outW > outH) { outH = Math.round(outH * (MAX / outW)); outW = MAX; }
    else { outW = Math.round(outW * (MAX / outH)); outH = MAX; }
  }

  const canvas = document.createElement('canvas');
  canvas.width = outW;
  canvas.height = outH;
  canvas.getContext('2d').drawImage(imageElement, 0, 0, outW, outH);

  return new Promise(resolve => {
    canvas.toBlob(blob => resolve({
      blob,
      dataUrl: canvas.toDataURL('image/jpeg', 0.92)
    }), 'image/jpeg', 0.92);
  });
}
