const express = require('express');
const mammoth = require('mammoth');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
  PageBreak, LevelFormat
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ─── TOM Brand constants ───────────────────────────────────────────────────
const BLUE  = "1AA1DB", GREEN = "38B18A", WHITE = "FFFFFF";
const NAVY  = "1F3864", LGRAY = "F2F2F2", LBLUE = "EAF5FB";
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 851;
const CONTENT_W = PAGE_W - MARGIN * 2;
const C1 = Math.floor(CONTENT_W / 3);
const C2 = Math.floor(CONTENT_W / 3);
const C3 = CONTENT_W - C1 - C2;

// ─── Style helpers ─────────────────────────────────────────────────────────
const thin  = (c="AAAAAA") => ({ style: BorderStyle.SINGLE, size: 2, color: c });
const thick = (c) => ({ style: BorderStyle.SINGLE, size: 6, color: c });
const none  = () => ({ style: BorderStyle.NONE, size: 0, color: WHITE });
const thinB = { top: thin(), bottom: thin(), left: thin(), right: thin() };
const noneB = { top: none(), bottom: none(), left: none(), right: none() };

const r = (text, opts={}) => new TextRun({ text, font:"Times New Roman", size:22, ...opts });
const p = (children, opts={}) => {
  if (typeof children === 'string') children = [r(children)];
  return new Paragraph({ children, spacing:{before:40,after:40}, ...opts });
};
const bp = (text, opts={}) => p([r(text,{bold:true})], opts);
const blankP = () => new Paragraph({ children:[r("")], spacing:{before:20,after:20} });

const hdrCell = (text, fill, w=C1) => new TableCell({
  borders:{top:thick(fill),bottom:thick(fill),left:thick(fill),right:thick(fill)},
  shading:{fill,type:ShadingType.CLEAR},
  margins:{top:100,bottom:100,left:130,right:130},
  verticalAlign:VerticalAlign.CENTER,
  width:{size:w,type:WidthType.DXA},
  children:[new Paragraph({
    children:[r(text,{bold:true,color:WHITE})],
    alignment:AlignmentType.CENTER,
    spacing:{before:60,after:60}
  })]
});

const secRow  = (en,ru,uz) => new TableRow({children:[hdrCell(en,BLUE),hdrCell(ru,NAVY),hdrCell(uz,GREEN)]});

const subSecRow = (en,ru,uz) => new TableRow({children:[en,ru,uz].map((txt,i)=>{
  const cols=[BLUE,NAVY,GREEN];
  return new TableCell({
    borders:thinB, shading:{fill:LBLUE,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:130,right:130}, verticalAlign:VerticalAlign.CENTER,
    width:{size:C1,type:WidthType.DXA},
    children:[p([r(txt,{bold:true,color:cols[i]})])]
  });
})});

const makeCell = (children, w=C1, shade=false) => {
  if (typeof children==='string') children=[p(children)];
  else if (children instanceof Paragraph) children=[children];
  else if (Array.isArray(children)) children=children.map(c=>typeof c==='string'?p(c):c);
  return new TableCell({
    borders:thinB, shading:{fill:shade?LGRAY:WHITE,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:130,right:130},
    verticalAlign:VerticalAlign.TOP, width:{size:w,type:WidthType.DXA}, children
  });
};

const triRow = (en,ru,uz,shade=false) => {
  const bg = shade?LGRAY:WHITE;
  const mc = (content) => {
    let ch;
    if (typeof content==='string') ch=[p(content)];
    else if (Array.isArray(content)) ch=content.map(c=>typeof c==='string'?p(c):c);
    else ch=[content];
    return new TableCell({
      borders:thinB, shading:{fill:bg,type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:130,right:130}, verticalAlign:VerticalAlign.TOP,
      width:{size:C1,type:WidthType.DXA}, children:ch
    });
  };
  return new TableRow({children:[mc(en),mc(ru),mc(uz)]});
};

const buildTable = (rows) => new Table({
  width:{size:CONTENT_W,type:WidthType.DXA},
  columnWidths:[C1,C2,C3],
  rows
});

const fieldRow = (label,en,ru,uz,shade=false) => {
  const bg=shade?LGRAY:WHITE;
  const fc=(val)=>new TableCell({
    borders:thinB, shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:70,bottom:70,left:130,right:130}, verticalAlign:VerticalAlign.TOP,
    width:{size:C1,type:WidthType.DXA},
    children:[new Paragraph({
      children:[r(label+":",{bold:true}),r("  "+val)],
      spacing:{before:40,after:40}
    })]
  });
  return new TableRow({children:[fc(en),fc(ru),fc(uz)]});
};

// ─── Bullet helper ─────────────────────────────────────────────────────────
const numbering = {
  config:[{
    reference:"bullets", levels:[{
      level:0, format:LevelFormat.BULLET,
      text:"\u2022", alignment:AlignmentType.LEFT,
      style:{paragraph:{indent:{left:360,hanging:180}}}
    }]
  }]
};

function bullet(text) {
  return new Paragraph({
    children:[r(text)],
    numbering:{reference:"bullets",level:0},
    spacing:{before:30,after:30}
  });
}

function bulletRow(enItems, ruItems, uzItems, shade=false) {
  const bg = shade?LGRAY:WHITE;
  const bc = (items) => new TableCell({
    borders:thinB, shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:80,right:130},
    verticalAlign:VerticalAlign.TOP,
    width:{size:C1,type:WidthType.DXA},
    children:items.map(i=>bullet(i))
  });
  return new TableRow({children:[bc(enItems),bc(ruItems),bc(uzItems)]});
}

// ─── Cover page ────────────────────────────────────────────────────────────
function coverPage(pos) {
  return [
    new Paragraph({
      children:[r("TALIMARJAN OPERATIONS AND MAINTENANCE LLC",{bold:true,size:28,color:WHITE})],
      alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
      shading:{fill:BLUE,type:ShadingType.CLEAR},
      border:{bottom:{style:BorderStyle.SINGLE,size:8,color:GREEN}}
    }),
    new Paragraph({
      children:[r("ООО «ТАЛИМАРДЖАН ОПЕРАШНС ЭНД МЕЙНТЕНАНС»",{bold:true,size:24,color:WHITE})],
      alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
      shading:{fill:NAVY,type:ShadingType.CLEAR}
    }),
    new Paragraph({
      children:[r("«TALIMARJAN OPERATIONS AND MAINTENANCE» MChJ",{bold:true,size:24,color:WHITE})],
      alignment:AlignmentType.CENTER, spacing:{before:0,after:120},
      shading:{fill:GREEN,type:ShadingType.CLEAR},
      border:{bottom:{style:BorderStyle.SINGLE,size:12,color:BLUE}}
    }),
    blankP(), blankP(), blankP(),
    new Paragraph({
      children:[r("JOB DESCRIPTION  |  ДОЛЖНОСТНАЯ ИНСТРУКЦИЯ  |  LAVOZIM TA'RIFI",{bold:true,size:26,color:NAVY})],
      alignment:AlignmentType.CENTER, spacing:{before:80,after:80}
    }),
    new Paragraph({
      children:[r("─────────────────────────────────────────────────────",{color:BLUE})],
      alignment:AlignmentType.CENTER, spacing:{before:0,after:60}
    }),
    new Table({
      width:{size:CONTENT_W,type:WidthType.DXA},
      columnWidths:[C1,C2,C3],
      rows:[new TableRow({children:[
        new TableCell({
          borders:{top:thick(BLUE),bottom:thick(BLUE),left:thick(BLUE),right:thin()},
          shading:{fill:LBLUE,type:ShadingType.CLEAR},
          margins:{top:120,bottom:120,left:160,right:160},
          verticalAlign:VerticalAlign.CENTER, width:{size:C1,type:WidthType.DXA},
          children:[
            new Paragraph({children:[r(pos.en,{bold:true,size:26,color:BLUE})],alignment:AlignmentType.CENTER,spacing:{before:40,after:20}}),
            new Paragraph({children:[r(pos.dept_en||"Chemical Water Treatment Workshop",{size:20,color:NAVY})],alignment:AlignmentType.CENTER,spacing:{before:0,after:40}})
          ]
        }),
        new TableCell({
          borders:{top:thick(NAVY),bottom:thick(NAVY),left:thin(),right:thin()},
          shading:{fill:"EEF0F5",type:ShadingType.CLEAR},
          margins:{top:120,bottom:120,left:160,right:160},
          verticalAlign:VerticalAlign.CENTER, width:{size:C2,type:WidthType.DXA},
          children:[
            new Paragraph({children:[r(pos.ru,{bold:true,size:26,color:NAVY})],alignment:AlignmentType.CENTER,spacing:{before:40,after:20}}),
            new Paragraph({children:[r(pos.dept_ru||"Цех химической водоподготовки",{size:20,color:NAVY})],alignment:AlignmentType.CENTER,spacing:{before:0,after:40}})
          ]
        }),
        new TableCell({
          borders:{top:thick(GREEN),bottom:thick(GREEN),left:thin(),right:thick(GREEN)},
          shading:{fill:"EAF7F3",type:ShadingType.CLEAR},
          margins:{top:120,bottom:120,left:160,right:160},
          verticalAlign:VerticalAlign.CENTER, width:{size:C3,type:WidthType.DXA},
          children:[
            new Paragraph({children:[r(pos.uz,{bold:true,size:26,color:GREEN})],alignment:AlignmentType.CENTER,spacing:{before:40,after:20}}),
            new Paragraph({children:[r(pos.dept_uz||"Suvni kimyoviy tozalash sexi",{size:20,color:"2E7D5B"})],alignment:AlignmentType.CENTER,spacing:{before:0,after:40}})
          ]
        })
      ]})]
    }),
    blankP(), blankP(),
    new Table({
      width:{size:CONTENT_W,type:WidthType.DXA},
      columnWidths:[Math.floor(CONTENT_W/2),CONTENT_W-Math.floor(CONTENT_W/2)],
      rows:[
        new TableRow({children:[
          makeCell([bp("Employee Copy  |  Экземпляр сотрудника  |  Xodim nusxasi")],Math.floor(CONTENT_W/2)),
          makeCell([p("Version 1.0  |  Версия 1.0  |  Versiya 1.0")],CONTENT_W-Math.floor(CONTENT_W/2))
        ]}),
        new TableRow({children:[
          makeCell([p("April 2025  |  Апрель 2025  |  Aprel 2025")],Math.floor(CONTENT_W/2)),
          makeCell([p("PRINT VERSION  |  Для печати  |  Chop etish uchun")],CONTENT_W-Math.floor(CONTENT_W/2))
        ]})
      ]
    }),
    blankP(), blankP(),
    new Table({
      width:{size:CONTENT_W,type:WidthType.DXA},
      columnWidths:[Math.floor(CONTENT_W*0.6),CONTENT_W-Math.floor(CONTENT_W*0.6)],
      rows:[new TableRow({children:[
        new TableCell({borders:noneB,width:{size:Math.floor(CONTENT_W*0.6),type:WidthType.DXA},children:[blankP()]}),
        new TableCell({
          borders:{top:thin(BLUE),bottom:thin(BLUE),left:thin(BLUE),right:thin(BLUE)},
          shading:{fill:LBLUE,type:ShadingType.CLEAR},
          margins:{top:100,bottom:100,left:160,right:160},
          width:{size:CONTENT_W-Math.floor(CONTENT_W*0.6),type:WidthType.DXA},
          children:[
            bp("APPROVED / УТВЕРЖДАЮ / TASDIQLANGAN"),
            p("General Director / Генеральный директор / Bosh direktor"),
            p("Talimarjan Operations and Maintenance LLC"),
            blankP(), p("Alaa Makki"), p("_________________________"),
            p("Date / Дата / Sana: ___________")
          ]
        })
      ]})]
    }),
    new Paragraph({children:[new PageBreak()]})
  ];
}

// ─── Welcome section ───────────────────────────────────────────────────────
function welcomeSection() {
  return [
    new Table({
      width:{size:CONTENT_W,type:WidthType.DXA},
      columnWidths:[C1,C2,C3],
      rows:[new TableRow({children:[
        new TableCell({borders:thinB,shading:{fill:LBLUE,type:ShadingType.CLEAR},margins:{top:120,bottom:120,left:160,right:160},width:{size:C1,type:WidthType.DXA},
          children:[
            p("Welcome to TOM. This document describes your role, responsibilities, what is expected of you, how your performance will be measured, and the path for your development at Talimarjan Operations and Maintenance LLC."),
            p("Please read it carefully and sign the acknowledgement on the last page.")
          ]}),
        new TableCell({borders:thinB,shading:{fill:"EEF0F5",type:ShadingType.CLEAR},margins:{top:120,bottom:120,left:160,right:160},width:{size:C2,type:WidthType.DXA},
          children:[
            p("Добро пожаловать в TOM. Настоящий документ описывает вашу роль, обязанности, требования к вам, порядок оценки результативности и возможные пути развития в ООО «Талимарджан Операшнс энд Мейнтенанс»."),
            p("Пожалуйста, прочитайте его внимательно и подпишите лист ознакомления на последней странице.")
          ]}),
        new TableCell({borders:thinB,shading:{fill:"EAF7F3",type:ShadingType.CLEAR},margins:{top:120,bottom:120,left:160,right:160},width:{size:C3,type:WidthType.DXA},
          children:[
            p("TOMga xush kelibsiz. Ushbu hujjat sizning lavozimingizdagi vazifalaringizni, sizdan nimalari kutilayotganini, ish samaradorligingiz qanday o'lchanishini va «Talimarjan Operations and Maintenance» MChJdagi rivojlanish yo'lingizni tavsiflaydi."),
            p("Iltimos, uni diqqat bilan o'qing va oxirgi sahifadagi tanishish qismini imzolang.")
          ]})
      ]})]
    }),
    blankP()
  ];
}

// ─── Acknowledgement ───────────────────────────────────────────────────────
function acknowledgementSection(pos, docCode, supervisor) {
  const supEn = supervisor.en || "Direct Supervisor";
  const supRu = supervisor.ru || "Непосредственный руководитель";
  const supUz = supervisor.uz || "Bevosita rahbar";
  return [
    buildTable([
      secRow("10. ACKNOWLEDGEMENT","10. ЛИСТ ОЗНАКОМЛЕНИЯ","10. TANISHIB CHIQILGANLIGI"),
      triRow(
        p("By signing below, you confirm that you have read and understood this Job Description. This document describes the role, not the individual. It may be updated periodically in line with business needs."),
        p("Подписывая настоящий документ, вы подтверждаете, что ознакомились с должностной инструкцией и поняли её содержание. Настоящий документ описывает должность, а не конкретного сотрудника. Он может периодически обновляться в соответствии с потребностями бизнеса."),
        p("Quyida imzo chekish orqali siz ushbu lavozim ta'rifini o'qib chiqib, tushunganingizni tasdiqlaysiz. Ushbu hujjat lavozimni tavsiflaydi, shaxsni emas. U biznes ehtiyojlariga muvofiq yangilanishi mumkin.")
      )
    ]),
    blankP(),
    new Table({
      width:{size:CONTENT_W,type:WidthType.DXA},
      columnWidths:[Math.floor(CONTENT_W/2),CONTENT_W-Math.floor(CONTENT_W/2)],
      rows:[new TableRow({children:[
        new TableCell({borders:thinB,margins:{top:120,bottom:120,left:160,right:160},width:{size:Math.floor(CONTENT_W/2),type:WidthType.DXA},
          children:[
            bp("Employee / Сотрудник / Xodim:"), blankP(),
            p("Full Name / Ф.И.О. / F.I.O.:"), p("_______________________________________"), blankP(),
            p("Signature / Подпись / Imzo:"), p("_______________________________________"), blankP(),
            p("Date / Дата / Sana:"), p("_______________________________________")
          ]}),
        new TableCell({borders:thinB,margins:{top:120,bottom:120,left:160,right:160},width:{size:CONTENT_W-Math.floor(CONTENT_W/2),type:WidthType.DXA},
          children:[
            bp("HR Director / HR Директор / HR Direktori:"), p("S. Puchka"), p("_______________________________________"), blankP(),
            bp(supEn+" / "+supRu+" / "+supUz+":"), p("_______________________________________"), blankP(),
            p("Date / Дата / Sana:"), p("_______________________________________")
          ]})
      ]})]
    }),
    blankP(),
    new Paragraph({children:[r("EMPLOYEE COPY — FOR PERSONAL USE ONLY. This document is the property of Talimarjan Operations and Maintenance LLC.",{size:16,color:"888888"})],alignment:AlignmentType.CENTER,spacing:{before:40,after:20}}),
    new Paragraph({children:[r(docCode+"  |  © TOM 2025",{size:16,color:"888888"})],alignment:AlignmentType.CENTER,spacing:{before:0,after:40}})
  ];
}

// ─── KPI table ─────────────────────────────────────────────────────────────
function kpiTable(kpiRows) {
  const nameW = Math.floor(CONTENT_W*0.27);
  const tgtW  = Math.floor(CONTENT_W*0.12);
  const frqW  = CONTENT_W - nameW*3 - tgtW;
  const kHdr  = () => new TableRow({children:["KPI (EN)","КПЭ (RU)","KPI (UZ)","Target / Цель","Frequency / Частота"].map((txt,i)=>{
    const fills=[BLUE,NAVY,GREEN,BLUE,BLUE];
    const ws=[nameW,nameW,nameW,tgtW,frqW];
    return new TableCell({
      borders:{top:thick(fills[i]),bottom:thick(fills[i]),left:thick(fills[i]),right:thick(fills[i])},
      shading:{fill:fills[i],type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:100,right:100},
      verticalAlign:VerticalAlign.CENTER, width:{size:ws[i],type:WidthType.DXA},
      children:[new Paragraph({children:[r(txt,{bold:true,color:WHITE})],alignment:AlignmentType.CENTER,spacing:{before:40,after:40}})]
    });
  })});
  const spanHdr = new TableRow({children:[new TableCell({
    borders:thinB, shading:{fill:BLUE,type:ShadingType.CLEAR},
    margins:{top:100,bottom:100,left:130,right:130}, columnSpan:5,
    children:[new Paragraph({children:[r("4. HOW YOUR PERFORMANCE WILL BE MEASURED  |  4. ОЦЕНКА РЕЗУЛЬТАТИВНОСТИ  |  4. SAMARADORLIK KO'RSATKICHLARI",{bold:true,color:WHITE})],alignment:AlignmentType.CENTER,spacing:{before:60,after:60}})]
  })]});
  const kRow=(en,ru,uz,tgt,frq,shade=false)=>{
    const bg=shade?LGRAY:WHITE; const ws=[nameW,nameW,nameW,tgtW,frqW];
    return new TableRow({children:[en,ru,uz,tgt,frq].map((txt,i)=>new TableCell({
      borders:thinB, shading:{fill:bg,type:ShadingType.CLEAR},
      margins:{top:70,bottom:70,left:100,right:100},
      verticalAlign:VerticalAlign.CENTER, width:{size:ws[i],type:WidthType.DXA},
      children:[p(txt)]
    }))});
  };
  return new Table({
    width:{size:CONTENT_W,type:WidthType.DXA},
    columnWidths:[nameW,nameW,nameW,tgtW,frqW],
    rows:[spanHdr, kHdr(), ...kpiRows.map((kr,i)=>kRow(kr[0],kr[1],kr[2],kr[3],kr[4],i%2===1))]
  });
}

// ─── Working conditions ────────────────────────────────────────────────────
function workingConditions(schedule) {
  return buildTable([
    secRow("6. WORKING CONDITIONS","6. УСЛОВИЯ ТРУДА","6. ISH SHAROITI"),
    triRow(
      [bp("Location:"),p("Talimarjan Power Plant, Kashkadarya Region, Republic of Uzbekistan.")],
      [bp("Местонахождение:"),p("Талимарджанская электростанция, Кашкадарьинская область, Республика Узбекистан.")],
      [bp("Joylashuv:"),p("Talimarjan elektr stansiyasi, Qashqadaryo viloyati, O'zbekiston Respublikasi.")]
    ),
    triRow(
      [bp("Work Schedule:"),p(schedule.en)],
      [bp("График работы:"),p(schedule.ru)],
      [bp("Ish jadvali:"),p(schedule.uz)],
      true
    ),
    triRow(
      [bp("HSE Requirements:"),p("All personnel must comply with TOM HSE standards, industrial safety regulations and Uzbek legislation. PPE must be worn at all times in the operational area.")],
      [bp("Требования ОТ и ТБ:"),p("Весь персонал обязан соблюдать стандарты TOM в области ОТ и ТБ, правила промышленной безопасности и законодательство РУз. В производственной зоне обязательно использование СИЗ.")],
      [bp("IST talablari:"),p("Barcha xodimlar TOM IST standartlari, sanoat xavfsizligi qoidalari va O'zbekiston qonunchiligiga rioya qilishi shart. Ishlab chiqarish zonasida doimo himoya vositalari kiyilishi shart.")]
    ),
    triRow(
      [bp("Travel:"),p("Not required.")],
      [bp("Командировки:"),p("Не требуются.")],
      [bp("Xizmat safari:"),p("Talab qilinmaydi.")],
      true
    )
  ]);
}

// ─── Career section ────────────────────────────────────────────────────────
function careerSection(upward, lateral) {
  return buildTable([
    secRow("9. YOUR CAREER DEVELOPMENT PATH AT TOM","9. КАРЬЕРНЫЙ ПУТЬ В TOM","9. TOMdagi MARTABA YO'LI"),
    triRow(
      p("TOM is committed to developing its people. Below are the development paths available to you with sustained high performance."),
      p("TOM привержен развитию своих сотрудников. При стабильно высоких результатах для вас открыты следующие пути развития."),
      p("TOM xodimlarini rivojlantirishga sodiqdir. Barqaror yuqori ko'rsatkichlar bilan quyidagi rivojlanish yo'llari sizga ochiq.")
    ),
    subSecRow("Upward Path","Карьерный рост","Yuqoriga yo'l"),
    triRow([bp(upward.en),p(upward.enDesc)],[bp(upward.ru),p(upward.ruDesc)],[bp(upward.uz),p(upward.uzDesc)],true),
    ...(lateral?[
      subSecRow("Lateral / Specialist Track","Горизонтальный / Экспертный путь","Gorizontal / Ekspert yo'l"),
      triRow([bp(lateral.en),p(lateral.enDesc)],[bp(lateral.ru),p(lateral.ruDesc)],[bp(lateral.uz),p(lateral.uzDesc)])
    ]:[]),
    subSecRow("What Will Help You Grow","Что поможет расти","O'sishingizga nima yordam beradi"),
    new TableRow({children:[["Deliver KPIs consistently above target for 2+ years","Lead at least one process improvement initiative","Demonstrate cross-functional collaboration","Complete relevant training programmes","Build strong relationships within your department and with related units"].map?(x=>x):null].filter(Boolean).length>0 ?
      [
        new TableCell({borders:thinB,shading:{fill:WHITE,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:80,right:130},verticalAlign:VerticalAlign.TOP,width:{size:C1,type:WidthType.DXA},
          children:["Deliver KPIs consistently above target for 2+ years","Lead at least one process improvement initiative","Demonstrate cross-functional collaboration","Complete relevant training programmes"].map(i=>new Paragraph({children:[new TextRun({text:i,font:"Times New Roman",size:22})],numbering:{reference:"bullets",level:0},spacing:{before:30,after:30}}))}),
        new TableCell({borders:thinB,shading:{fill:WHITE,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:80,right:130},verticalAlign:VerticalAlign.TOP,width:{size:C2,type:WidthType.DXA},
          children:["Стабильно выполнять КПЭ выше целевых показателей на протяжении 2+ лет","Возглавить хотя бы одну инициативу по совершенствованию процессов","Демонстрировать межфункциональное взаимодействие","Пройти соответствующие программы обучения"].map(i=>new Paragraph({children:[new TextRun({text:i,font:"Times New Roman",size:22})],numbering:{reference:"bullets",level:0},spacing:{before:30,after:30}}))}),
        new TableCell({borders:thinB,shading:{fill:WHITE,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:80,right:130},verticalAlign:VerticalAlign.TOP,width:{size:C3,type:WidthType.DXA},
          children:["2+ yil davomida KPIlarni maqsaddan yuqori bajarib borish","Kamida bitta jarayon takomillashtirish tashabbusiga rahbarlik qilish","Funksiyalararo hamkorlikni namoyish etish","Tegishli o'quv dasturlarini yakunlash"].map(i=>new Paragraph({children:[new TextRun({text:i,font:"Times New Roman",size:22})],numbering:{reference:"bullets",level:0},spacing:{before:30,after:30}}))})
      ] : []
    })
  ]);
}

// ─── Main document builder ─────────────────────────────────────────────────
async function buildJD(data) {
  const { pos, docCode, reportsTo, directReports, grade, schedule,
          duties, kpis, education, expYears, languages,
          career, supervisor, purpose } = data;

  const children = [
    ...coverPage(pos),
    ...welcomeSection(),

    // Section 1
    buildTable([
      secRow("1. YOUR ROLE AT A GLANCE","1. КРАТКАЯ ИНФОРМАЦИЯ О ДОЛЖНОСТИ","1. LAVOZIM HAQIDA QISQACHA"),
      fieldRow("Job Title", pos.en, pos.ru, pos.uz),
      fieldRow("Department", pos.dept_en||"", pos.dept_ru||"", pos.dept_uz||"", true),
      fieldRow("Reports To", reportsTo.en, reportsTo.ru, reportsTo.uz),
      fieldRow("Direct Reports", directReports, directReports, directReports, true),
      fieldRow("HAY Grade", grade, grade, grade),
      fieldRow("Job Code", docCode, docCode, docCode, true),
      fieldRow("Version / Date", "V1.0 | April 2025", "V1.0 | Апрель 2025", "V1.0 | Aprel 2025"),
    ]),
    blankP(),

    // Section 2
    buildTable([
      secRow("2. WHY THIS ROLE EXISTS","2. ЦЕЛЬ ДОЛЖНОСТИ","2. LAVOZIM MAQSADI"),
      triRow(p(purpose.en), p(purpose.ru), p(purpose.uz))
    ]),
    blankP(),

    // Section 3
    buildTable([
      secRow("3. KEY RESPONSIBILITIES","3. ОСНОВНЫЕ ОБЯЗАННОСТИ","3. ASOSIY MAJBURIYATLAR"),
      ...duties.flatMap((d,i)=>[
        subSecRow(d.title.en, d.title.ru, d.title.uz),
        bulletRow(d.en, d.ru, d.uz, i%2===1)
      ])
    ]),
    blankP(),

    // Section 4 KPIs
    kpiTable(kpis),
    blankP(),

    // Section 5
    buildTable([
      secRow("5. WHAT WE EXPECT FROM YOU","5. ТРЕБОВАНИЯ К КАНДИДАТУ","5. SIZDAN NIMA KUTAMIZ"),
      subSecRow("Education","Образование","Ta'lim"),
      triRow(p(education.en), p(education.ru), p(education.uz), true),
      subSecRow("Experience","Опыт","Tajriba"),
      triRow(
        p(`Minimum ${expYears} year(s) of practical experience in a relevant role.`),
        p(`Минимум ${expYears} год(а) практического опыта на аналогичной должности.`),
        p(`Tegishli lavozimda kamida ${expYears} yillik amaliy tajriba.`)
      ),
      subSecRow("Languages & Software","Языки и ПО","Tillar va dasturiy ta'minot"),
      triRow(p(languages.en), p(languages.ru), p(languages.uz), true)
    ]),
    blankP(),

    // Section 6
    workingConditions(schedule),
    blankP(),

    // Section 7
    buildTable([
      secRow("7. YOUR RIGHTS & AUTHORITY","7. ВАШИ ПРАВА И ПОЛНОМОЧИЯ","7. SIZNING HUQUQ VA VAKOLATINGIZ"),
      bulletRow(
        ["Perform duties within the scope defined in this Job Description","Request information and resources from your direct supervisor to complete assigned tasks","Raise safety concerns or equipment defects immediately to your supervisor","Propose improvements to work methods and procedures"],
        ["Выполнять обязанности в рамках настоящей должностной инструкции","Запрашивать у непосредственного руководителя информацию и ресурсы для выполнения задач","Незамедлительно сообщать руководителю о проблемах безопасности или дефектах оборудования","Вносить предложения по улучшению методов работы и процедур"],
        ["Ushbu lavozim ta'rifida belgilangan doirada vazifalarni bajarish","Berilgan vazifalarni bajarish uchun zarur ma'lumot va resurslarni bevosita rahbaringizdan talab qilish","Xavfsizlik muammolari yoki uskunalar nuqsonlarini darhol rahbaringizga bildirish","Ish usullari va tartiblarini takomillashtirish bo'yicha takliflar kiritish"]
      )
    ]),
    blankP(),

    // Section 8 First 90 days
    buildTable([
      secRow("8. YOUR FIRST 90 DAYS","8. ПЕРВЫЕ 90 ДНЕЙ","8. BIRINCHI 90 KUN"),
      subSecRow("Days 1–30: Learn","Дни 1–30: Знакомство","1–30 kun: O'rganish"),
      bulletRow(
        ["Complete mandatory safety induction and HSE training","Study all applicable instructions and operational regulations","Shadow your predecessor or a senior colleague for the first two weeks"],
        ["Пройти обязательный вводный инструктаж и обучение по ОТ и ТБ","Изучить все применимые инструкции и правила эксплуатации","Провести первые две недели с наставником или опытным коллегой"],
        ["Majburiy xavfsizlik yo'riqnomasi va IST o'qitishini yakunlash","Barcha amaldagi yo'riqnomalar va ekspluatatsiya qoidalarini o'rganish","Birinchi ikki haftani nastavnik yoki tajribali hamkasb bilan o'tkazish"]
      ),
      subSecRow("Days 31–60: Engage","Дни 31–60: Включение","31–60 kun: Faol ishtirok"),
      bulletRow(
        ["Work independently under periodic supervision","Demonstrate knowledge through practical checks","Raise any questions or knowledge gaps to your supervisor"],
        ["Работать самостоятельно под периодическим контролем руководителя","Демонстрировать знания в ходе практических проверок","Сообщать руководителю о вопросах и пробелах в знаниях"],
        ["Bevosita rahbaringizning davriy nazorati ostida mustaqil ishlash","Amaliy tekshiruvlar orqali bilimlaringizni ko'rsatish","Savol yoki bilimlaringizdagi bo'shliqlarni rahbaringizga bildirish"],
        true
      ),
      subSecRow("Days 61–90: Confirm","Дни 61–90: Подтверждение","61–90 kun: Tasdiqlash"),
      bulletRow(
        ["Pass mandatory qualification checks for independent work","Confirm personal equipment assignments and workplace inventory","Discuss 90-day observations and improvement ideas with your supervisor"],
        ["Пройти обязательные квалификационные проверки для самостоятельной работы","Подтвердить закреплённое оборудование и инвентарь рабочего места","Обсудить с руководителем итоги первых 90 дней"],
        ["Mustaqil ish uchun majburiy malaka tekshiruvlaridan o'tish","Shaxsiy uskunalar va ish joyi inventarini tasdiqlash","90 kunlik xulosalar va takomillashtirish g'oyalarini rahbaringiz bilan muhokama qilish"]
      )
    ]),
    blankP(),

    // Section 9 Career
    careerSection(career.upward, career.lateral||null),
    blankP(),

    // Section 10 Acknowledgement
    ...acknowledgementSection(pos, docCode, supervisor)
  ];

  const doc = new Document({
    numbering,
    styles:{
      default:{
        document:{
          run:{font:"Times New Roman",size:22},
          paragraph:{spacing:{before:40,after:40}}
        }
      }
    },
    sections:[{
      properties:{
        page:{
          size:{width:PAGE_W,height:PAGE_H},
          margin:{top:MARGIN,bottom:MARGIN,left:MARGIN,right:MARGIN}
        }
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

// ─── Routes ────────────────────────────────────────────────────────────────

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'TOM docx-service', version: '1.2.0' });
});

// Claude parse: POST /claude-parse
// Body: { text: '<anketa text>', filename: '<name>' }
// Returns: structured JSON extracted by Claude
app.post('/claude-parse', async (req, res) => {
  try {
    const { text, filename, apiKey } = req.body;
    if (!text) return res.status(400).json({ error: 'Missing text field' });
    if (!apiKey) return res.status(400).json({ error: 'Missing apiKey field' });

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 4096,
        system: `You are an expert HR analyst at TOM LLC (Talimarjan Operations and Maintenance LLC), a 900MW CCGT power plant in Uzbekistan.
Extract structured data from a job questionnaire written in Uzbek.
Return ONLY a valid JSON object with NO additional text, markdown, or explanation.
Use this exact structure:
{
  "position_uz": "<full position name>",
  "department_uz": "<department name>",
  "reports_to_uz": "<direct supervisor>",
  "has_subordinates": true/false,
  "subordinate_count": 0,
  "duties": [{"description_uz": "<duty>", "importance": 10, "time_pct": 20}],
  "substitute_for_uz": "<who this substitutes>",
  "substituted_by_uz": "<who substitutes this>",
  "education_level": "secondary|secondary_special|higher_nonspec|higher_spec",
  "special_skills_uz": "<skills or Yo'q>",
  "experience_years": 3,
  "management_experience_required": false,
  "language_level": "none|basic|intermediate|advanced",
  "software_level": "none|standard|advanced",
  "managerial_authority": false,
  "special_conditions_uz": "<conditions or Shart emas>"
}`,
        messages: [{ role: 'user', content: 'Extract structured data:\n\n' + text }]
      })
    });

    const data = await response.json();
    if (!response.ok) return res.status(response.status).json(data);

    const content = data.content[0].text;
    const clean = content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const parsed = JSON.parse(clean);
    parsed.source_filename = filename || 'unknown.docx';
    res.json(parsed);
  } catch (err) {
    console.error('[claude-parse]', err);
    res.status(500).json({ error: err.message });
  }
});

// Claude generate JD: POST /claude-generate
// Body: { structuredData: {...}, apiKey: 'sk-ant-...' }
// Returns: complete JD JSON for /generate endpoint
app.post('/claude-generate', async (req, res) => {
  try {
    const { structuredData, apiKey } = req.body;
    if (!structuredData) return res.status(400).json({ error: 'Missing structuredData' });
    if (!apiKey) return res.status(400).json({ error: 'Missing apiKey' });

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 8192,
        system: `You are a senior HR specialist at TOM LLC (Talimarjan Operations and Maintenance LLC), a 900MW CCGT power plant in Uzbekistan.
Write professional trilingual Job Descriptions (EN | RU | UZ).
CEO: Alaa Makki. HR Director: S. Puchka. Location: Talimarjan Power Plant, Kashkadarya Region.
Return ONLY a valid JSON object with NO additional text.
Required structure:
{
  "pos": {"en":"","ru":"","uz":"","dept_en":"","dept_ru":"","dept_uz":""},
  "docCode": "TOM_INT_HRD_JD_[DEPT]_[SEQ]",
  "filename": "TOM_INT_HRD_JD_[DEPT]_[SEQ]_EMP_[Title]_EN_RU_UZ_V1.docx",
  "reportsTo": {"en":"","ru":"","uz":""},
  "directReports": "None / Нет / Yo'q",
  "grade": "GSO-X | Level",
  "purpose": {"en":"","ru":"","uz":""},
  "duties": [{"title":{"en":"3.1 Title (X%)","ru":"","uz":""},"en":[],"ru":[],"uz":[]}],
  "kpis": [["EN name","RU name","UZ name","Target","Frequency"]],
  "education": {"en":"","ru":"","uz":""},
  "expYears": 3,
  "languages": {"en":"","ru":"","uz":""},
  "schedule": {"en":"","ru":"","uz":""},
  "career": {"upward":{"en":"","enDesc":"","ru":"","ruDesc":"","uz":"","uzDesc":""},"lateral":null},
  "supervisor": {"en":"","ru":"","uz":""}
}`,
        messages: [{
          role: 'user',
          content: 'Generate complete trilingual JD for this position:\n\n' + JSON.stringify(structuredData, null, 2) + '\n\nRules:\n- duties time% must sum to 100\n- Generate 5 relevant KPIs\n- Shift workers: 12-hour rotating schedule\n- Office workers: 5-day week\n- Use power plant terminology'
        }]
      })
    });

    const data = await response.json();
    if (!response.ok) return res.status(response.status).json(data);

    const content = data.content[0].text;
    const clean = content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jdData = JSON.parse(clean);
    jdData.source_filename = structuredData.source_filename;
    res.json(jdData);
  } catch (err) {
    console.error('[claude-generate]', err);
    res.status(500).json({ error: err.message });
  }
});

// Main endpoint: POST /generate
// Body: JD data JSON (same structure as genPosition cfg)
app.post('/generate', async (req, res) => {
  const start = Date.now();
  try {
    const data = req.body;
    if (!data.pos || !data.docCode) {
      return res.status(400).json({ error: 'Missing required fields: pos, docCode' });
    }
    const buffer = await buildJD(data);
    const filename = data.filename || `${data.docCode}_EN_RU_UZ_V1.docx`;
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'X-Generation-Ms': Date.now() - start
    });
    res.send(buffer);
  } catch (err) {
    console.error('[generate]', err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

// Parse endpoint: POST /parse
// Body: { base64: '<base64 docx content>' }
// Returns: structured JSON extracted from questionnaire
// All-in-one: POST /process
// Body: { base64: '<docx>', apiKey: 'sk-ant-...' }
// Returns: complete JD JSON (parse + generate in one call)
app.post('/process', async (req, res) => {
  try {
    const { base64, apiKey, filename } = req.body;
    if (!base64) return res.status(400).json({ error: 'Missing base64' });
    if (!apiKey) return res.status(400).json({ error: 'Missing apiKey' });

    // Step 1: Extract text from docx
    const buffer = Buffer.from(base64, 'base64');
    const extracted = await mammoth.extractRawText({ buffer });
    const text = (extracted.value || '').substring(0, 8000);

    if (!text.trim()) return res.status(400).json({ error: 'Empty document text' });

    // Step 2: Single Claude call - parse AND generate in one prompt
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 8192,
        system: `You are a senior HR specialist at TOM LLC (Talimarjan Operations and Maintenance LLC), a 900MW CCGT power plant in Uzbekistan.
You will receive a job questionnaire in Uzbek. In ONE response, extract the data AND generate a complete trilingual Job Description.
CEO: Alaa Makki. HR Director: S. Puchka. Location: Talimarjan Power Plant, Kashkadarya Region, Uzbekistan.

Return ONLY a valid JSON object with NO additional text, markdown, or explanation:
{
  "pos": {"en":"","ru":"","uz":"","dept_en":"","dept_ru":"","dept_uz":""},
  "docCode": "TOM_INT_HRD_JD_[DEPT]_[SEQ]",
  "filename": "TOM_INT_HRD_JD_[DEPT]_[SEQ]_EMP_[Title]_EN_RU_UZ_V1.docx",
  "reportsTo": {"en":"","ru":"","uz":""},
  "directReports": "None / Нет / Yo'q",
  "grade": "GSO-X | Level",
  "purpose": {"en":"","ru":"","uz":""},
  "duties": [{"title":{"en":"3.1 Title (X%)","ru":"","uz":""},"en":[],"ru":[],"uz":[]}],
  "kpis": [["EN name","RU name","UZ name","Target","Frequency"]],
  "education": {"en":"","ru":"","uz":""},
  "expYears": 3,
  "languages": {"en":"","ru":"","uz":""},
  "schedule": {"en":"","ru":"","uz":""},
  "career": {"upward":{"en":"","enDesc":"","ru":"","ruDesc":"","uz":"","uzDesc":""},"lateral":null},
  "supervisor": {"en":"","ru":"","uz":""}
}
Rules: duties time% must sum to 100. Generate 5 KPIs. Shift workers get 12-hour rotating schedule. Office workers get 5-day week. Use power plant terminology.`,
        messages: [{
          role: 'user',
          content: `Generate complete trilingual JD from this questionnaire:\n\n${text}`
        }]
      })
    });

    const data = await response.json();
    if (!response.ok) return res.status(response.status).json(data);

    const content = data.content[0].text;
    const clean = content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jdData = JSON.parse(clean);
    jdData.source_filename = filename || 'unknown.docx';
    res.json(jdData);
  } catch (err) {
    console.error('[process]', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/parse', async (req, res) => {
  try {
    const { base64 } = req.body;
    if (!base64) return res.status(400).json({ error: 'Missing base64 field' });
    const buffer = Buffer.from(base64, 'base64');
    const result = await mammoth.extractRawText({ buffer });
    res.json({ text: result.value, messages: result.messages });
  } catch (err) {
    console.error('[parse]', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`TOM docx-service listening on port ${PORT}`);
});
