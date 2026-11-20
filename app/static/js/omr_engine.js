// Pure Front-end Answer Sheet OMR Engine for Browser Execution
// Coordinates mirror engine/config.py (A4 size: 595 x 842 pt)

const CONFIG = {
  PAGE_W_PT: 595,
  PAGE_H_PT: 842,
  CORNER_MARK_MARGIN: 18,
  CORNER_MARK_SIZE: 24,

  LEFT_X: 45,
  RIGHT_X: 595 - 45,
  NAME_LINE_X0: 45 + 40,

  GRADE_VALUES: [7, 8, 9, 10, 11, 12],
  GRADE_W: 15.0,
  GRADE_H: 7.5,
  GRADE_CIRCLE_Y: 842 - 174,
  GRADE_STEP_X: 26,

  CLASS_VALUES: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
  CLASS_W: 14.0,
  CLASS_H: 7.5,
  CLASS_CIRCLE_Y: 842 - 199,
  CLASS_STEP_X: 21,

  SEAT_BOX_W: 200,
  SEAT_BOX_H: 62,
  SEAT_BOX_X0: (595 - 45) - 200,
  SEAT_BOX_Y0: 842 - 205,
  SEAT_W: 12.5,
  SEAT_H: 7.0,
  SEAT_STEP_X: 16,
  SEAT_START_X: ((595 - 45) - 200) + 38,
  SEAT_TOP_ROW_Y: (842 - 205) + 40,
  SEAT_BOTTOM_ROW_Y: (842 - 205) + 18,

  DIVIDER_Y: 842 - 220,

  BOX_X0: 25,
  BOX_X1: 595 - 25,
  BOX_Y0: 55,
  BOX_Y1: (842 - 220) - 12,
  BOX_PAD: 10,

  BUBBLE_W: 13.0,
  BUBBLE_H: 7.5,
  COL_COUNT: 3,
  MAX_QUESTIONS: 100,
  Q_NUM_AREA_W: 30,
  GAP_AFTER_NUM: 10,

  MIN_SCORE_GRADE: 0.05,
  MIN_SCORE_CLASS: 0.05,
  MIN_SCORE_SEAT: 0.04,
  MIN_SCORE_CHOICE: 0.04,
};

function makeChoices(count) {
  const choices = [];
  for (let i = 0; i < count; i++) {
    choices.push(String.fromCharCode(65 + i));
  }
  return choices;
}

function computeAnswerLayout(numQuestions, choicesCount = 4) {
  const choices = makeChoices(choicesCount);
  const boxW = CONFIG.BOX_X1 - CONFIG.BOX_X0;
  const colW = boxW / CONFIG.COL_COUNT;
  const rowsPerCol = Math.ceil(CONFIG.MAX_QUESTIONS / CONFIG.COL_COUNT);

  const innerY0 = CONFIG.BOX_Y0 + CONFIG.BOX_PAD;
  const innerY1 = CONFIG.BOX_Y1 - CONFIG.BOX_PAD;

  const firstRowY = innerY1 - 30;
  const bottomLimit = innerY0 + CONFIG.BUBBLE_H / 2;
  const usableHeight = firstRowY - bottomLimit;
  const rowStep = rowsPerCol > 1 ? usableHeight / (rowsPerCol - 1) : 0;

  const bubbleXs = [];
  for (let c = 0; c < CONFIG.COL_COUNT; c++) {
    const cx0 = CONFIG.BOX_X0 + c * colW;
    const innerLeft = cx0 + CONFIG.BOX_PAD;
    const innerRight = cx0 + colW - CONFIG.BOX_PAD;

    const qRight = innerLeft + CONFIG.Q_NUM_AREA_W;
    const xA = qRight + CONFIG.GAP_AFTER_NUM + CONFIG.BUBBLE_W / 2;
    const xLast = innerRight - CONFIG.BUBBLE_W / 2;
    const step = (xLast - xA) / (choices.length - 1);
    const xs = [];
    for (let i = 0; i < choices.length; i++) {
      xs.push(xA + i * step);
    }
    bubbleXs.push(xs);
  }

  return {
    rowsPerCol,
    firstRowY,
    rowStep,
    bubbleXs,
  };
}

class OMREngine {
  static warpCanvas(canvas) {
    if (typeof cv === 'undefined') return canvas;
    const src = cv.imread(canvas);
    const gray = new cv.Mat();
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY, 0);
    const h = gray.rows;
    const w = gray.cols;
    const zoom = Math.min(w / CONFIG.PAGE_W_PT, h / CONFIG.PAGE_H_PT);
    const warped = OMREngine._warpToCanonical(src, gray, zoom);
    const outCanvas = document.createElement("canvas");
    cv.imshow(outCanvas, warped);
    src.delete();
    gray.delete();
    warped.delete();
    return outCanvas;
  }

  static async processCanvas(canvas, numQuestions = 50, choicesCount = 4) {
    if (typeof cv === 'undefined') {
      throw new Error("OpenCV.js is not loaded");
    }

    const src = cv.imread(canvas);
    const gray = new cv.Mat();
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY, 0);

    const h = gray.rows;
    const w = gray.cols;
    const zoom = Math.min(w / CONFIG.PAGE_W_PT, h / CONFIG.PAGE_H_PT);

    // Corner Mark Detection & Perspective Transformation
    const warped = OMREngine._warpToCanonical(src, gray, zoom);

    // Score extraction from warped canonical A4 canvas
    const result = OMREngine._analyzeCanonical(warped, zoom, numQuestions, choicesCount);

    src.delete();
    gray.delete();
    warped.delete();

    return result;
  }

  static _warpToCanonical(src, gray, zoom) {
    const h = gray.rows;
    const w = gray.cols;

    // Corner ROI Search
    const rx = Math.floor(w * 0.25);
    const ry = Math.floor(h * 0.25);

    const corners = OMREngine._findCorners(gray, zoom, rx, ry, w, h);

    const canonicalW = Math.round(CONFIG.PAGE_W_PT * zoom);
    const canonicalH = Math.round(CONFIG.PAGE_H_PT * zoom);

    const dstMat = new cv.Mat(canonicalH, canonicalW, src.type());

    if (!corners || corners.length < 4) {
      // Fallback: simple resize
      cv.resize(src, dstMat, new cv.Size(canonicalW, canonicalH));
      return dstMat;
    }

    const srcTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
      corners[0].x, corners[0].y,
      corners[1].x, corners[1].y,
      corners[2].x, corners[2].y,
      corners[3].x, corners[3].y
    ]);

    const mMargin = CONFIG.CORNER_MARK_MARGIN + CONFIG.CORNER_MARK_SIZE / 2.0;
    const dstTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
      mMargin * zoom, mMargin * zoom,
      (CONFIG.PAGE_W_PT - mMargin) * zoom, mMargin * zoom,
      (CONFIG.PAGE_W_PT - mMargin) * zoom, (CONFIG.PAGE_H_PT - mMargin) * zoom,
      mMargin * zoom, (CONFIG.PAGE_H_PT - mMargin) * zoom
    ]);

    const M = cv.getPerspectiveTransform(srcTri, dstTri);
    cv.warpPerspective(src, dstMat, M, new cv.Size(canonicalW, canonicalH));

    srcTri.delete();
    dstTri.delete();
    M.delete();

    return dstMat;
  }

  static _findCorners(gray, zoom, rx, ry, w, h) {
    // simplified robust corner detector
    const expectedCenter = (CONFIG.CORNER_MARK_MARGIN + CONFIG.CORNER_MARK_SIZE / 2.0) * zoom;
    const expectedPx = CONFIG.CORNER_MARK_SIZE * zoom;

    const findCornerInRoi = (roi, targetX, targetY, xOff, yOff) => {
      const blur = new cv.Mat();
      cv.GaussianBlur(roi, blur, new cv.Size(5, 5), 0);
      const bw = new cv.Mat();
      cv.threshold(blur, bw, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);

      const contours = new cv.MatVector();
      const hierarchy = new cv.Mat();
      cv.findContours(bw, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);

      let best = null;
      let maxScore = -1;

      for (let i = 0; i < contours.size(); ++i) {
        const cnt = contours.get(i);
        const area = cv.contourArea(cnt);
        if (area < (expectedPx * expectedPx) * 0.15) continue;

        const rect = cv.boundingRect(cnt);
        const ar = rect.width / (rect.height + 1e-6);
        if (ar < 0.6 || ar > 1.4) continue;

        const cx = rect.x + rect.width / 2.0;
        const cy = rect.y + rect.height / 2.0;
        const dx = cx - targetX;
        const dy = cy - targetY;
        const dist = Math.sqrt(dx * dx + dy * dy);

        const fill = area / (rect.width * rect.height + 1e-6);
        if (fill < 0.5) continue;

        const score = (area * fill) / (1.0 + dist);
        if (score > maxScore) {
          maxScore = score;
          best = { x: cx + xOff, y: cy + yOff };
        }
      }

      blur.delete();
      bw.delete();
      contours.delete();
      hierarchy.delete();

      return best;
    };

    const tlRoi = gray.roi(new cv.Rect(0, 0, rx, ry));
    const trRoi = gray.roi(new cv.Rect(w - rx, 0, rx, ry));
    const blRoi = gray.roi(new cv.Rect(0, h - ry, rx, ry));
    const brRoi = gray.roi(new cv.Rect(w - rx, h - ry, rx, ry));

    let tl = findCornerInRoi(tlRoi, expectedCenter, expectedCenter, 0, 0);
    let tr = findCornerInRoi(trRoi, rx - expectedCenter, expectedCenter, w - rx, 0);
    let bl = findCornerInRoi(blRoi, expectedCenter, ry - expectedCenter, 0, h - ry);
    let br = findCornerInRoi(brRoi, rx - expectedCenter, ry - expectedCenter, w - rx, h - ry);

    tlRoi.delete();
    trRoi.delete();
    blRoi.delete();
    brRoi.delete();

    // Parallelogram completion for missing 4th corner (tl + br = tr + bl)
    if (!tl && tr && bl && br) {
      tl = { x: tr.x + bl.x - br.x, y: tr.y + bl.y - br.y };
    } else if (!tr && tl && bl && br) {
      tr = { x: tl.x + br.x - bl.x, y: tl.y + br.y - bl.y };
    } else if (!br && tl && tr && bl) {
      br = { x: tr.x + bl.x - tl.x, y: tr.y + bl.y - tl.y };
    } else if (!bl && tl && tr && br) {
      bl = { x: tl.x + br.x - tr.x, y: tl.y + br.y - tr.y };
    }

    if (tl && tr && bl && br) {
      return [tl, tr, br, bl];
    }
    return null;
  }

  static _analyzeCanonical(warped, zoom, numQuestions, choicesCount) {
    const gray = new cv.Mat();
    cv.cvtColor(warped, gray, cv.COLOR_RGBA2GRAY, 0);

    const bw = new cv.Mat();
    cv.threshold(gray, bw, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);

    // Differential scoring robust to printed borders & text:
    // Measures how much darker the inner core is vs. pure background paper.
    // By discarding dark printed border/text pixels from the background sample,
    // it prevents (bg - inner) from becoming negative on small/dense bubbles.
    const getBubbleScore = (ptX, ptY, wPt, hPt) => {
      const cx = Math.round(ptX * zoom);
      const cy = Math.round((CONFIG.PAGE_H_PT - ptY) * zoom); // PDF y-up to image y-down
      const halfW = Math.max(2, (wPt * zoom) / 2.0);
      const halfH = Math.max(2, (hPt * zoom) / 2.0);

      // Inner core rectangle (60% of half dimensions)
      const coreHalfW = halfW * 0.60;
      const coreHalfH = halfH * 0.60;

      // Background rectangle margin (110% to 165% of half dimensions)
      const bgMinHalfW = halfW * 1.10;
      const bgMaxHalfW = halfW * 1.65;
      const bgMinHalfH = halfH * 1.10;
      const bgMaxHalfH = halfH * 1.65;

      const innerVals = [];
      const bgVals = [];

      const y0 = Math.max(0, Math.floor(cy - bgMaxHalfH));
      const y1 = Math.min(gray.rows - 1, Math.ceil(cy + bgMaxHalfH));
      const x0 = Math.max(0, Math.floor(cx - bgMaxHalfW));
      const x1 = Math.min(gray.cols - 1, Math.ceil(cx + bgMaxHalfW));

      for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
          const dx = Math.abs(x - cx);
          const dy = Math.abs(y - cy);
          const pixVal = gray.ucharPtr(y, x)[0]; // grayscale 0-255

          if (dx <= coreHalfW && dy <= coreHalfH) {
            innerVals.push(pixVal);
          } else if (
            (dx <= bgMaxHalfW && dy <= bgMaxHalfH) &&
            (dx >= bgMinHalfW || dy >= bgMinHalfH)
          ) {
            bgVals.push(pixVal);
          }
        }
      }

      if (innerVals.length < 5 || bgVals.length < 5) return 0;

      // Filter out dark printed borders & numbers by taking the upper 50% brightest background pixels
      bgVals.sort((a, b) => a - b);
      const topHalfBg = bgVals.slice(Math.floor(bgVals.length / 2));
      let bgSum = 0;
      for (let i = 0; i < topHalfBg.length; i++) bgSum += topHalfBg[i];
      const bgMean = bgSum / topHalfBg.length;

      let innerSum = 0;
      for (let i = 0; i < innerVals.length; i++) innerSum += innerVals[i];
      const innerMean = innerSum / innerVals.length;

      // Score = how much darker inner is compared to pure background paper
      let score = (bgMean - innerMean) / 255.0;
      if (score < 0) score = 0;
      if (score > 1) score = 1;
      return score;
    };

    // Grade Detection
    let bestGrade = null;
    let maxGradeScore = CONFIG.MIN_SCORE_GRADE;
    const gradeXs = [CONFIG.NAME_LINE_X0, CONFIG.NAME_LINE_X0 + 26, CONFIG.NAME_LINE_X0 + 52, CONFIG.NAME_LINE_X0 + 78, CONFIG.NAME_LINE_X0 + 104, CONFIG.NAME_LINE_X0 + 130];
    CONFIG.GRADE_VALUES.forEach((val, idx) => {
      const score = getBubbleScore(gradeXs[idx], CONFIG.GRADE_CIRCLE_Y, CONFIG.GRADE_W, CONFIG.GRADE_H);
      if (score > maxGradeScore) {
        maxGradeScore = score;
        bestGrade = val;
      }
    });

    // Class Detection
    let bestClass = null;
    let maxClassScore = CONFIG.MIN_SCORE_CLASS;
    const classXs = Array.from({ length: 10 }, (_, i) => CONFIG.NAME_LINE_X0 + i * CONFIG.CLASS_STEP_X);
    CONFIG.CLASS_VALUES.forEach((val, idx) => {
      const score = getBubbleScore(classXs[idx], CONFIG.CLASS_CIRCLE_Y, CONFIG.CLASS_W, CONFIG.CLASS_H);
      if (score > maxClassScore) {
        maxClassScore = score;
        bestClass = val;
      }
    });

    // Seat Detection
    const tensScores = [];
    const unitsScores = [];
    for (let d = 0; d < 10; d++) {
      const x = CONFIG.SEAT_START_X + d * CONFIG.SEAT_STEP_X;
      tensScores.push(getBubbleScore(x, CONFIG.SEAT_TOP_ROW_Y, CONFIG.SEAT_W, CONFIG.SEAT_H));
      unitsScores.push(getBubbleScore(x, CONFIG.SEAT_BOTTOM_ROW_Y, CONFIG.SEAT_W, CONFIG.SEAT_H));
    }

    const tensMax = Math.max(...tensScores);
    const tensDigit = tensMax >= CONFIG.MIN_SCORE_SEAT ? tensScores.indexOf(tensMax) : null;

    const unitsMax = Math.max(...unitsScores);
    const unitsDigit = unitsMax >= CONFIG.MIN_SCORE_SEAT ? unitsScores.indexOf(unitsMax) : null;

    let seatNo = "";
    if (tensDigit !== null && unitsDigit !== null) {
      seatNo = `${tensDigit}${unitsDigit}`;
    } else if (unitsDigit !== null && tensDigit === null) {
      // 1-9 號學生若僅在個位數劃記（十位數留白），自動補 0（例如個位劃 7 -> 辨識為 07 號）
      seatNo = `0${unitsDigit}`;
    } else {
      // 若僅劃記十位數但個位數漏劃（如只劃十位 1、個位空白），或全空：屬劃記不完整，絕不擅自猜號或補 0，設為空白供人工核對
      seatNo = "";
    }

    // Answers Detection
    const layout = computeAnswerLayout(numQuestions, choicesCount);
    const choices = makeChoices(choicesCount);
    const answers = [];

    for (let q = 1; q <= numQuestions; q++) {
      const colIdx = Math.floor((q - 1) / layout.rowsPerCol);
      const rowInCol = (q - 1) % layout.rowsPerCol;
      const y = layout.firstRowY - rowInCol * layout.rowStep;
      const xs = layout.bubbleXs[colIdx];

      const qScores = [];
      for (let cIdx = 0; cIdx < choices.length; cIdx++) {
        qScores.push(getBubbleScore(xs[cIdx], y, CONFIG.BUBBLE_W, CONFIG.BUBBLE_H));
      }

      const bestScore = Math.max(...qScores);
      if (bestScore < CONFIG.MIN_SCORE_CHOICE) {
        answers.push("");
      } else {
        // Multi-choice picking: pick all choices with score >= max(MIN_SCORE_CHOICE, bestScore * 0.65)
        const multiCut = Math.max(CONFIG.MIN_SCORE_CHOICE, bestScore * 0.65);
        const picked = [];
        for (let cIdx = 0; cIdx < choices.length; cIdx++) {
          if (qScores[cIdx] >= multiCut) {
            picked.push(choices[cIdx]);
          }
        }
        answers.push(picked.join(""));
      }
    }

    gray.delete();
    bw.delete();

    return {
      grade: bestGrade || "",
      class: bestClass !== null ? bestClass : "",
      seat: seatNo,
      answers,
    };
  }
}
