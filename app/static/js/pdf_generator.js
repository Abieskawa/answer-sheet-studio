// Pure Frontend Canvas-based Answer Sheet PDF Generator (High-DPI 1:1 A4 Exact Match)
// Matches engine/generator.py coordinates precisely

async function generateAnswerSheetPDF(titleText = "定期評量 答案卷", subject = "", numQuestions = 50, choicesCount = 4) {
  if (typeof PDFLib === 'undefined') {
    throw new Error("pdf-lib is not loaded");
  }

  // Exact A4 Dimensions in points (1/72 inch)
  const PAGE_W_PT = 595.28;
  const PAGE_H_PT = 841.89;

  // Render scale factor for crisp printing (300 DPI equivalent)
  const SCALE = 3.0;

  const canvas = document.createElement("canvas");
  canvas.width = Math.round(PAGE_W_PT * SCALE);
  canvas.height = Math.round(PAGE_H_PT * SCALE);
  const ctx = canvas.getContext("2d");

  // White background
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Helper coordinate mapper (Pt -> HighDPI Canvas Px, and Y-flip for PDF Cartesian -> Canvas Top-Down)
  const ptX = (xPt) => xPt * SCALE;
  const ptY = (yPt) => (PAGE_H_PT - yPt) * SCALE;
  const ptS = (valPt) => valPt * SCALE;

  ctx.fillStyle = "#000000";
  ctx.strokeStyle = "#000000";

  // 1. Corner Alignment Squares (MARGIN=18pt, SIZE=24pt)
  const m = 18;
  const s = 24;
  ctx.fillRect(ptX(m), ptY(PAGE_H_PT - m), ptS(s), ptS(s));
  ctx.fillRect(ptX(PAGE_W_PT - m - s), ptY(PAGE_H_PT - m), ptS(s), ptS(s));
  ctx.fillRect(ptX(m), ptY(m + s), ptS(s), ptS(s));
  ctx.fillRect(ptX(PAGE_W_PT - m - s), ptY(m + s), ptS(s), ptS(s));

  // 2. Header
  const titleY = PAGE_H_PT - 70;
  const subtitleY = PAGE_H_PT - 95;
  const headerLineY = PAGE_H_PT - 110;
  const boxX0 = 25;
  const boxX1 = PAGE_W_PT - 25;

  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.font = `bold ${ptS(20)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
  ctx.fillText(titleText || "定期評量 答案卷", ptX(PAGE_W_PT / 2), ptY(titleY));

  if (subject) {
    ctx.font = `${ptS(13)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
    ctx.fillText(`科目：${subject}`, ptX(PAGE_W_PT / 2), ptY(subtitleY));
  }

  ctx.lineWidth = ptS(1.0);
  ctx.beginPath();
  ctx.moveTo(ptX(boxX0), ptY(headerLineY));
  ctx.lineTo(ptX(boxX1), ptY(headerLineY));
  ctx.stroke();

  // 3. Identity Area
  const leftX = 45;
  const rightX = PAGE_W_PT - 45;
  const nameY = PAGE_H_PT - 145;
  const nameLineX0 = leftX + 40;
  const nameLineX1 = leftX + 195;

  ctx.textAlign = "left";
  ctx.font = `${ptS(12)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
  ctx.fillText("姓名", ptX(leftX), ptY(nameY));

  ctx.lineWidth = ptS(0.9);
  ctx.beginPath();
  ctx.moveTo(ptX(nameLineX0), ptY(nameY - 2));
  ctx.lineTo(ptX(nameLineX1), ptY(nameY - 2));
  ctx.stroke();

  // Grade Rectangles
  const gradeY = PAGE_H_PT - 174;
  const gradeValues = [7, 8, 9, 10, 11, 12];
  const gradeW = 15.0, gradeH = 7.5;
  ctx.fillText("年級", ptX(leftX), ptY(gradeY - 2));

  ctx.textAlign = "center";
  ctx.font = `${ptS(9)}px sans-serif`;
  gradeValues.forEach((v, i) => {
    const x = nameLineX0 + i * 26;
    ctx.fillText(String(v), ptX(x), ptY(gradeY + 11));
    ctx.strokeRect(ptX(x - gradeW / 2), ptY(gradeY + gradeH / 2), ptS(gradeW), ptS(gradeH));
  });

  // Class Rectangles
  const classY = PAGE_H_PT - 199;
  const classValues = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
  const classW = 14.0, classH = 7.5;
  ctx.textAlign = "left";
  ctx.font = `${ptS(12)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
  ctx.fillText("班級", ptX(leftX), ptY(classY - 2));

  ctx.textAlign = "center";
  ctx.font = `${ptS(9)}px sans-serif`;
  classValues.forEach((v, i) => {
    const x = nameLineX0 + i * 21;
    ctx.fillText(String(v), ptX(x), ptY(classY + 11));
    ctx.strokeRect(ptX(x - classW / 2), ptY(classY + classH / 2), ptS(classW), ptS(classH));
  });

  // Seat Box
  const seatBoxW = 200;
  const seatBoxH = 62;
  const seatBoxX0 = rightX - seatBoxW;
  const seatBoxY0 = PAGE_H_PT - 205;
  const seatW = 12.5, seatH = 7.0;
  const seatStartXs = seatBoxX0 + 38;
  const seatTopY = seatBoxY0 + 40;
  const seatBottomY = seatBoxY0 + 18;

  ctx.lineWidth = ptS(1.2);
  ctx.strokeRect(ptX(seatBoxX0), ptY(seatBoxY0 + seatBoxH), ptS(seatBoxW), ptS(seatBoxH));

  ctx.font = `${ptS(11)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
  ctx.fillText("座", ptX(seatBoxX0 - 26), ptY(seatBoxY0 + 38));
  ctx.fillText("號", ptX(seatBoxX0 - 26), ptY(seatBoxY0 + 18));

  ctx.font = `${ptS(9)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
  ctx.fillText("十位", ptX(seatBoxX0 + 14), ptY(seatTopY - 3));
  ctx.fillText("個位", ptX(seatBoxX0 + 14), ptY(seatBottomY - 3));

  ctx.font = `${ptS(8)}px sans-serif`;
  for (let d = 0; d < 10; d++) {
    const x = seatStartXs + d * 16;
    ctx.fillText(String(d), ptX(x), ptY(seatTopY + 11));
    ctx.strokeRect(ptX(x - seatW / 2), ptY(seatTopY + seatH / 2), ptS(seatW), ptS(seatH));

    ctx.fillText(String(d), ptX(x), ptY(seatBottomY + 10));
    ctx.strokeRect(ptX(x - seatW / 2), ptY(seatBottomY + seatH / 2), ptS(seatW), ptS(seatH));
  }

  // Divider
  const dividerY = PAGE_H_PT - 220;
  ctx.lineWidth = ptS(1.0);
  ctx.beginPath();
  ctx.moveTo(ptX(boxX0), ptY(dividerY));
  ctx.lineTo(ptX(boxX1), ptY(dividerY));
  ctx.stroke();

  // 4. Answer Outer Box
  const boxY1 = dividerY - 12;
  const boxY0 = 55;
  ctx.lineWidth = ptS(2.0);
  ctx.strokeRect(ptX(boxX0), ptY(boxY1), ptS(boxX1 - boxX0), ptS(boxY1 - boxY0));

  // 5. Questions Layout (3 Columns)
  const cols = 3;
  const choices = Array.from({ length: choicesCount }, (_, i) => String.fromCharCode(65 + i));
  const boxW = boxX1 - boxX0;
  const colW = boxW / cols;
  const rowsPerCol = Math.ceil(100 / cols);
  const usedCols = Math.min(cols, Math.ceil(numQuestions / rowsPerCol));

  const boxPad = 10;
  const innerY0 = boxY0 + boxPad;
  const innerY1 = boxY1 - boxPad;
  const headerY = innerY1 - 10;
  const firstRowY = innerY1 - 30;
  const bubbleW = 13.0, bubbleH = 7.5;
  const bottomLimit = innerY0 + bubbleH / 2;
  const usable = firstRowY - bottomLimit;
  const rowStep = rowsPerCol > 1 ? usable / (rowsPerCol - 1) : 0;

  const bubbleXs = [];
  const qNumRightX = [];

  for (let c = 0; c < cols; c++) {
    const cx0 = boxX0 + c * colW;
    const innerLeft = cx0 + boxPad;
    const innerRight = cx0 + colW - boxPad;

    const qRight = innerLeft + 30;
    const xA = qRight + 10 + bubbleW / 2;
    const xLast = innerRight - bubbleW / 2;
    const step = (xLast - xA) / (choices.length - 1);
    const xs = [];
    for (let i = 0; i < choices.length; i++) {
      xs.push(xA + i * step);
    }
    bubbleXs.push(xs);
    qNumRightX.push(qRight);
  }

  // Draw Headers A, B, C, D
  ctx.font = `${ptS(10)}px sans-serif`;
  for (let col = 0; col < usedCols; col++) {
    const xs = bubbleXs[col];
    choices.forEach((ch, j) => {
      ctx.fillText(ch, ptX(xs[j]), ptY(headerY));
    });
  }

  // Draw Question Rows & Rectangular Bubbles
  ctx.lineWidth = ptS(1.1);
  let q = 1;
  for (let col = 0; col < usedCols; col++) {
    for (let r = 0; r < rowsPerCol; r++) {
      if (q > numQuestions) break;
      const y = firstRowY - r * rowStep;
      const qx = qNumRightX[col];

      ctx.textAlign = "right";
      ctx.font = `${ptS(9)}px sans-serif`;
      ctx.fillText(`${q}.`, ptX(qx), ptY(y - 3));

      const xs = bubbleXs[col];
      xs.forEach((x) => {
        ctx.strokeRect(ptX(x - bubbleW / 2), ptY(y + bubbleH / 2), ptS(bubbleW), ptS(bubbleH));
      });
      q++;
    }
  }

  // Footer text
  ctx.textAlign = "center";
  ctx.font = `${ptS(11)}px "Microsoft JhengHei", "PingFang TC", sans-serif`;
  ctx.fillText("請使用 2B 鉛筆或深色原子筆將方格塗滿", ptX(PAGE_W_PT / 2), ptY(42));

  // Convert Canvas to High-Res Image Data URL -> Insert into 1:1 A4 PDF Document
  const imgDataUrl = canvas.toDataURL("image/png");
  const { PDFDocument } = PDFLib;
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([PAGE_W_PT, PAGE_H_PT]);

  const pngImage = await pdfDoc.embedPng(imgDataUrl);
  page.drawImage(pngImage, {
    x: 0,
    y: 0,
    width: PAGE_W_PT,
    height: PAGE_H_PT,
  });

  const pdfBytes = await pdfDoc.save();
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${titleText || "answer_sheet"}.pdf`;
  link.click();
}
