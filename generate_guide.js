const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, ColumnBreak
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function heading1(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 28, color: "1F3864", font: "Calibri" })],
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 4 } },
  });
}

function heading2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, color: "1F3864", font: "Calibri" })],
    spacing: { before: 200, after: 80 },
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: "Calibri", ...opts })],
    spacing: { before: 60, after: 60 },
  });
}

function bodyBold(text) {
  return body(text, { bold: true });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: [new TextRun({ text, size: 22, font: "Calibri" })],
    spacing: { before: 40, after: 40 },
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    children: [new TextRun({ text, size: 22, font: "Calibri" })],
    spacing: { before: 40, after: 40 },
  });
}

function spacer(before = 80, after = 80) {
  return new Paragraph({ children: [new TextRun("")], spacing: { before, after } });
}

// Placeholder box: gray table with centered italic text
function placeholder(description) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            width: { size: 9026, type: WidthType.DXA },
            shading: { fill: "EBEBEB", type: ShadingType.CLEAR },
            margins: { top: 360, bottom: 360, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 120, after: 120 },
                children: [new TextRun({ text: description, italics: true, size: 20, color: "666666", font: "Calibri" })]
              })
            ]
          })
        ]
      })
    ]
  });
}

// Code block: single-cell table with Courier New
function codeBlock(lines) {
  const codeParas = lines.map(line =>
    new Paragraph({
      children: [new TextRun({ text: line, font: "Courier New", size: 20, color: "333333" })],
      spacing: { before: 40, after: 40 },
    })
  );
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder },
            width: { size: 9026, type: WidthType.DXA },
            shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: codeParas
          })
        ]
      })
    ]
  });
}

// ── Palette table ─────────────────────────────────────────────────────────────
const palette = [
  { type: "0",  name: "BLACK",  hex: "000000" },
  { type: "1",  name: "BLUE",   hex: "1246A2" },
  { type: "2",  name: "BROWN",  hex: "8C4914" },
  { type: "3",  name: "CYAN",   hex: "45D8E9" },
  { type: "4",  name: "GREEN",  hex: "2DCA2D" },
  { type: "5",  name: "ORANGE", hex: "F88716" },
  { type: "6",  name: "PINK",   hex: "F380DE" },
  { type: "7",  name: "PURPLE", hex: "9738F1" },
  { type: "8",  name: "RED",    hex: "BB1416" },
  { type: "9",  name: "WHITE",  hex: "ffffff" },
  { type: "10", name: "YELLOW", hex: "FFE000" },
];

function paletteTable() {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      makeHeaderCell("Type", 1000),
      makeHeaderCell("Tên", 2500),
      makeHeaderCell("Màu hex", 3000),
      makeHeaderCell("Swatch", 2526),
    ]
  });

  const dataRows = palette.map(p => new TableRow({
    children: [
      makeDataCell(p.type, 1000),
      makeDataCell(p.name, 2500),
      makeDataCell("#" + p.hex, 3000),
      makeSwatchCell(p.hex, 2526),
    ]
  }));

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1000, 2500, 3000, 2526],
    rows: [headerRow, ...dataRows]
  });
}

function makeHeaderCell(text, w) {
  return new TableCell({
    borders: cellBorders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: true, size: 20, font: "Calibri" })]
    })]
  });
}

function makeDataCell(text, w) {
  return new TableCell({
    borders: cellBorders,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, size: 20, font: "Calibri" })]
    })]
  });
}

function makeSwatchCell(hex, w) {
  // Show a small colored rectangle using table shading
  return new TableCell({
    borders: cellBorders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: hex.toUpperCase() === "FFFFFF" ? "FFFFFF" : hex.toUpperCase(), type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text: "#" + hex, size: 18, font: "Calibri",
        color: isLight(hex) ? "333333" : "FFFFFF" })]
    })]
  });
}

function isLight(hex) {
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  return (r * 0.299 + g * 0.587 + b * 0.114) > 128;
}

// ── Build document ────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Calibri" })
          ]
        })]
      })
    },
    children: [
      // ── TITLE ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 320 },
        children: [new TextRun({
          text: "Hướng Dẫn Sử Dụng Công Cụ",
          bold: true, size: 40, font: "Calibri", color: "1F3864"
        })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({
          text: "Pixel Art \u2192 Game Level JSON",
          bold: true, size: 40, font: "Calibri", color: "2E74B5"
        })]
      }),

      // ── TỔNG QUAN ──
      heading1("TỔNG QUAN"),
      spacer(80, 60),
      body("Bộ công cụ gồm 2 file HTML:"),
      bullet("GenJson.html \u2013 Đọc màu từ ảnh pixel art, map sang type number, xuất JSON levelDataPixel"),
      bullet("LoadJson.html \u2013 Load JSON đó, hiển thị bản đồ, chỉnh sửa màu từng ô, gom nhóm màu, xuất lại JSON"),
      spacer(100, 80),
      bodyBold("Bảng màu chuẩn (ETypeColor):"),
      spacer(60, 60),
      paletteTable(),
      spacer(120, 60),

      // ── PHẦN 1 ──
      heading1("PHẦN 1: GENJSON.HTML \u2013 ĐỌC PIXEL ART"),
      spacer(80, 60),

      heading2("Bước 1: Mở file và upload ảnh"),
      body("Mở GenJson.html bằng trình duyệt. Nhấn nút \"Upload Image\" hoặc kéo thả file ảnh pixel art vào trang. Ảnh sẽ hiển thị trên canvas."),
      spacer(60, 60),
      placeholder("Minh họa giao diện GenJson với ảnh pixel art đã upload lên canvas, có nút Upload Image ở góc trên"),
      spacer(120, 60),

      heading2("Bước 2: Chọn chế độ chia lưới"),
      body("Có 2 chế độ:"),
      spacer(60, 40),

      bodyBold("Chế độ A \u2013 Pixel Grid (mặc định):"),
      body("Nhập trực tiếp kích thước ô:"),
      bullet("W (Block Width): chiều rộng mỗi ô pixel (mặc định: 10)"),
      bullet("H (Block Height): chiều cao mỗi ô pixel (mặc định: 10)"),
      bullet("OX (Offset X): dịch lưới theo chiều ngang"),
      bullet("OY (Offset Y): dịch lưới theo chiều dọc"),
      body("Phù hợp khi biết chính xác kích thước từng ô."),
      spacer(60, 60),
      placeholder("Minh họa ô nhập W=10, H=10, OX=0, OY=0 và lưới kẻ đều trên ảnh"),
      spacer(120, 60),

      bodyBold("Chế độ B \u2013 Mark Mode (chấm tay):"),
      body("Chọn tab \"Mark\". Click chuột lên ảnh để đặt dấu chấm:"),
      bullet("Click vào cột \u2192 tạo đường kẻ dọc"),
      bullet("Click vào hàng \u2192 tạo đường kẻ ngang"),
      body("Công cụ sẽ kẻ lưới theo đúng vị trí các chấm đã đánh. Dùng khi ảnh có ô không đều hoặc có viền/padding giữa các ô."),
      spacer(60, 60),
      placeholder("Minh họa các chấm đỏ/xanh được click trên ảnh, và các đường lưới tự động kẻ theo"),
      spacer(120, 60),

      heading2("Bước 3: Lấy màu từ tâm ô"),
      body("Sau khi chia lưới, công cụ tự động lấy màu pixel ở trung tâm mỗi ô. Màu đó sẽ được so sánh với bảng PALETTE để tìm type gần nhất (dùng thuật toán CIELAB Delta-E cho độ chính xác cao)."),
      spacer(80, 60),

      heading2("Bước 4: Xuất JSON"),
      body("Nhấn \"Generate JSON\". File JSON sẽ hiển thị trong ô text bên dưới:"),
      spacer(60, 60),
      codeBlock([
        "{",
        '  "cols": 10,',
        '  "rows": 8,',
        '  "colorMapping": { "8": "#BB1416", "9": "#ffffff" },',
        '  "cells": [',
        '    { "x": 0, "y": 0, "type": 9 },',
        '    { "x": 1, "y": 0, "type": 8 }',
        "  ]",
        "}",
      ]),
      spacer(100, 60),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [
          new TextRun({ text: "Lưu ý tọa độ: ", bold: true, size: 22, font: "Calibri" }),
          new TextRun({ text: "Gốc (0,0) nằm ở góc dưới bên trái. Trục Y tăng dần lên trên.", size: 22, font: "Calibri" }),
        ]
      }),
      body("Nhấn \"Copy\" để copy JSON vào clipboard, sau đó dán vào file game."),
      spacer(120, 60),

      // ── PHẦN 2 ──
      heading1("PHẦN 2: LOADJSON.HTML \u2013 XEM VÀ CHỈNH SỬA BẢN ĐỒ"),
      spacer(80, 60),

      heading2("Bước 1: Load JSON"),
      body("Mở LoadJson.html. Dán nội dung JSON vào ô text input rồi nhấn \"Load\". Bản đồ sẽ được vẽ ra canvas với đúng màu sắc theo colorMapping."),
      spacer(60, 60),
      placeholder("Minh họa paste JSON vào textbox và bản đồ render ra với màu đúng"),
      spacer(120, 60),

      heading2("Bước 2: Xem thông tin ô"),
      body("Di chuột qua từng ô trên bản đồ để xem tooltip hiển thị:"),
      bullet("Tọa độ (x, y)"),
      bullet("Type của ô"),
      bullet("Màu hex hiện tại"),
      spacer(60, 60),
      placeholder("Minh họa tooltip nổi lên khi hover ô, ví dụ: (3,2) type 8 #BB1416"),
      spacer(120, 60),

      heading2("Bước 3: Sửa màu từng ô"),
      body("Click vào ô bất kỳ trên bản đồ \u2192 một color picker nổi lên. Chọn màu mới rồi nhấn Apply. Ô đó sẽ đổi màu ngay lập tức. Có thể sửa nhiều ô khác nhau trước khi lưu."),
      spacer(60, 60),
      placeholder("Minh họa color picker nổi lên sau khi click ô, với nút Apply và Cancel"),
      spacer(120, 60),

      heading2("Bước 4: Gom nhóm màu (Group Colors)"),
      body("Sau khi chỉnh sửa xong, nhấn \"Group Colors\". Công cụ sẽ:"),
      numbered("Gom các ô có màu gần nhau thành cùng 1 nhóm"),
      numbered("Tự động map mỗi nhóm sang type PALETTE gần nhất"),
      numbered("Cập nhật toàn bộ bản đồ theo type mới"),
      spacer(60, 40),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [
          new TextRun({ text: "Ví dụ: ", bold: true, size: 22, font: "Calibri" }),
          new TextRun({ text: "nhiều ô màu trắng hơi khác nhau (#fffffe, #f5f5f5, #ffffff) sẽ được gom thành type 9 (WHITE = #ffffff).", size: 22, font: "Calibri" }),
        ]
      }),
      spacer(60, 60),
      placeholder("Minh họa trước/sau khi nhấn Group Colors - bản đồ trở nên gọn màu hơn"),
      spacer(120, 60),

      heading2("Bước 5: Xuất JSON"),
      body("Nhấn \"Export JSON\". JSON mới sẽ hiển thị với type đã được cập nhật và colorMapping chính xác theo bảng PALETTE."),
      spacer(120, 60),

      // ── LƯU Ý ──
      heading1("LƯU Ý QUAN TRỌNG"),
      spacer(80, 60),

      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "Tọa độ (0,0) = góc dưới trái", bold: true, size: 22, font: "Calibri" }),
          new TextRun({ text: " \u2013 Khác với canvas HTML thông thường (góc trên trái). Đây là tọa độ game.", size: 22, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "colorMapping", bold: true, size: 22, font: "Courier New" }),
          new TextRun({ text: " \u2013 Lưu mapping type \u2192 hex để LoadJson có thể khôi phục màu chính xác khi reload.", size: 22, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "CIELAB Delta-E", bold: true, size: 22, font: "Calibri" }),
          new TextRun({ text: " \u2013 Thuật toán so màu dựa trên cảm nhận thị giác, cho kết quả chính xác hơn RGB thông thường.", size: 22, font: "Calibri" }),
        ]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "Default block size = 10", bold: true, size: 22, font: "Calibri" }),
          new TextRun({ text: " \u2013 Để tránh lag khi ảnh lớn. Giảm xuống nếu ô pixel nhỏ hơn.", size: 22, font: "Calibri" }),
        ]
      }),
      spacer(120, 60),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("G:/ReadPixel/HuongDan.docx", buffer);
  console.log("Done: G:/ReadPixel/HuongDan.docx (" + buffer.length + " bytes)");
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
