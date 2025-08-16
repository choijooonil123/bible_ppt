// ===== 1) 성경 권 이름 매핑 (정식명 + 약칭들) =====
// 'canon' = 파일명(bible/폴더)에서 사용할 정식 이름
const BOOKS = [
  { canon:"창세기", aliases:["창"] },
  { canon:"출애굽기", aliases:["출"] },
  { canon:"레위기", aliases:["레"] },
  { canon:"민수기", aliases:["민"] },
  { canon:"신명기", aliases:["신"] },
  { canon:"여호수아", aliases:["수"] },
  { canon:"사사기", aliases:["삿"] },
  { canon:"룻기", aliases:["룻"] },
  { canon:"사무엘상", aliases:["삼상"] },
  { canon:"사무엘하", aliases:["삼하"] },
  { canon:"열왕기상", aliases:["왕상"] },
  { canon:"열왕기하", aliases:["왕하"] },
  { canon:"역대상", aliases:["대상"] },
  { canon:"역대하", aliases:["대하"] },
  { canon:"에스라", aliases:["스"] },
  { canon:"느헤미야", aliases:["느"] },
  { canon:"에스더", aliases:["에"] },
  { canon:"욥기", aliases:["욥"] },
  { canon:"시편", aliases:["시"] },
  { canon:"잠언", aliases:["잠"] },
  { canon:"전도서", aliases:["전"] },
  { canon:"아가", aliases:["아"] },
  { canon:"이사야", aliases:["사"] },
  { canon:"예레미야", aliases:["렘"] },
  { canon:"예레미야애가", aliases:["애"] },
  { canon:"에스겔", aliases:["겔"] },
  { canon:"다니엘", aliases:["단"] },
  { canon:"호세아", aliases:["호"] },
  { canon:"요엘", aliases:["욜"] },
  { canon:"아모스", aliases:["암"] },
  { canon:"오바댜", aliases:["옵"] },
  { canon:"요나", aliases:["욘"] },
  { canon:"미가", aliases:["미"] },
  { canon:"나훔", aliases:["나"] },
  { canon:"하박국", aliases:["합"] },
  { canon:"스바냐", aliases:["습"] },
  { canon:"학개", aliases:["학"] },
  { canon:"스가랴", aliases:["슥"] },
  { canon:"말라기", aliases:["말"] },

  { canon:"마태복음", aliases:["마"] },
  { canon:"마가복음", aliases:["막"] },
  { canon:"누가복음", aliases:["눅"] },
  { canon:"요한복음", aliases:["요"] },
  { canon:"사도행전", aliases:["행"] },
  { canon:"로마서", aliases:["롬"] },
  { canon:"고린도전서", aliases:["고전"] },
  { canon:"고린도후서", aliases:["고후"] },
  { canon:"갈라디아서", aliases:["갈"] },
  { canon:"에베소서", aliases:["엡"] },
  { canon:"빌립보서", aliases:["빌"] },
  { canon:"골로새서", aliases:["골"] },
  { canon:"데살로니가전서", aliases:["살전"] },
  { canon:"데살로니가후서", aliases:["살후"] },
  { canon:"디모데전서", aliases:["딤전"] },
  { canon:"디모데후서", aliases:["딤후"] },
  { canon:"디도서", aliases:["딛"] },
  { canon:"빌레몬서", aliases:["몬"] },
  { canon:"히브리서", aliases:["히"] },
  { canon:"야고보", aliases:["약"] },
  { canon:"베드로전서", aliases:["벧전"] },
  { canon:"베드로후서", aliases:["벧후"] },
  { canon:"요한일서", aliases:["요일"] },
  { canon:"요한이서", aliases:["요이"] },
  { canon:"요한삼서", aliases:["요삼"] },
  { canon:"유다서", aliases:["유"] },
  { canon:"요한계시록", aliases:["계"] },
];

const BOOK_ALIAS_TO_CANON = (() => {
  const map = new Map();
  BOOKS.forEach(({canon, aliases}) => {
    map.set(canon, canon);
    aliases.forEach(a => map.set(a, canon));
  });
  return map;
})();

function getAllBookKeys() {
  const arr = [];
  BOOKS.forEach(({canon, aliases}) => {
    arr.push(canon, ...aliases);
  });
  return arr;
}

// ===== 2) 기존 DOM 바인딩 (index.html 그대로 사용) =====
const refEl   = document.getElementById("ref");
const bookEl  = document.getElementById("book");
const dlEl    = document.getElementById("bookList");
const addBtn  = document.getElementById("addRef");
const goBtn   = document.getElementById("go");

const bgFileEl   = document.getElementById("bgFile");
const blurPxEl   = document.getElementById("blurPx");
const maxCharsEl = document.getElementById("maxChars");
const baseFontEl = document.getElementById("baseFont");
const bodyColorEl  = document.getElementById("bodyColor");
const titleColorEl = document.getElementById("titleColor");

const showVerseNoEl     = document.getElementById("showVerseNo");
const oneVersePerSlideEl= document.getElementById("oneVersePerSlide");
const makeTitleSlideEl  = document.getElementById("makeTitleSlide");

// 자동완성 목록: 정식명 + 약칭
getAllBookKeys().forEach(name => {
  const opt = document.createElement("option");
  opt.value = name;
  dlEl.appendChild(opt);
});

// ===== 3) 유틸 =====
function trimAll(s){ return (s||"").replace(/\s+/g," ").trim(); }
function splitToItems(input){
  return input.split(/[\n,]+/).map(s=>trimAll(s)).filter(Boolean);
}

function normalizeBookName(inputName){
  const key = trimAll(inputName);
  const canon = BOOK_ALIAS_TO_CANON.get(key);
  if (!canon) {
    throw new Error(`알 수 없는 성경 이름: "${inputName}"`);
  }
  return canon;
}

// 입력 파서: 세 가지 형식 지원
//  A) "요한복음 3:16~4:3"  (다른 장 포함)
//  B) "요한복음 3:16~18"   (같은 장, 끝절만)
//  C) "요한복음 3:16"      (단일 절)
function parseRef(text){
  const t = trimAll(text);

  const pOtherCh = /^(.+?)\s+(\d+):(\d+)\s*~\s*(\d+):(\d+)$/; // A
  const pSameCh  = /^(.+?)\s+(\d+):(\d+)\s*~\s*(\d+)$/;       // B
  const pSingle  = /^(.+?)\s+(\d+):(\d+)$/;                   // C

  let m = t.match(pOtherCh);
  if (m) {
    const book = normalizeBookName(m[1]);
    return { book, sCh:+m[2], sV:+m[3], eCh:+m[4], eV:+m[5] };
  }
  m = t.match(pSameCh);
  if (m) {
    const book = normalizeBookName(m[1]);
    return { book, sCh:+m[2], sV:+m[3], eCh:+m[2], eV:+m[4] };
  }
  m = t.match(pSingle);
  if (m) {
    const book = normalizeBookName(m[1]);
    return { book, sCh:+m[2], sV:+m[3], eCh:+m[2], eV:+m[3] }; // 단일 절
  }
  throw new Error(`입력 형식 오류: "${text}"
허용 예) "요 3:16", "요 3:16~18", "요 3:16~4:3"`);
}

async function fetchVerses(book, sCh, sV, eCh, eV) {
  const res = await fetch(`bible/${book}.txt`);
  if (!res.ok) throw new Error(`성경 파일을 찾을 수 없습니다: bible/${book}.txt`);
  const txt = await res.text();
  const lines = txt.split(/\r?\n/);
  const out = [];
  for (const line of lines) {
    const m = line.match(/^(\d+):(\d+)\s+(.*)$/); // "장:절 내용"
    if (!m) continue;
    const ch = +m[1], v = +m[2], body = m[3].trim();
    if ((ch > eCh) || (ch === eCh && v > eV)) break;
    if ((ch > sCh) || (ch === sCh && v >= sV)) {
      out.push({ch, v, text: body});
    }
  }
  return out;
}

function splitByLen(s, maxLen){
  const arr = [];
  for (let i=0;i<s.length;i+=maxLen) arr.push(s.slice(i, i+maxLen));
  return arr;
}
function autoFont(base, totalChars){
  if (totalChars <= 40) return base;
  const over = totalChars - 40;
  const dec = Math.min(14, Math.floor(over * 0.3));
  return Math.max(18, base - dec);
}
function linesForVerse(verseText, maxLen){ return splitByLen(verseText, maxLen); }

function blobToDataURL(blob){
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}

async function imgToDataURLWithBlur(fileOrUrl, blurPx=0){
  let imgURL;
  if (fileOrUrl instanceof File) {
    imgURL = URL.createObjectURL(fileOrUrl);
  } else {
    imgURL = fileOrUrl;
  }
  const img = await new Promise((res, rej) => {
    const im = new Image();
    im.crossOrigin = "anonymous"; // same-origin 권장
    im.onload = () => res(im);
    im.onerror = rej;
    im.src = imgURL;
  });
  const canvas = document.createElement("canvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  const ctx = canvas.getContext("2d");
  if (blurPx > 0) ctx.filter = `blur(${blurPx}px)`;
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/jpeg", 0.92);
}

// ===== 4) UI: 보조 버튼/생성 버튼 =====
addBtn.addEventListener("click", () => {
  const b = trimAll(bookEl.value);
  if (!b) return;
  if (refEl.value && !/[\n,]\s*$/.test(refEl.value)) refEl.value += ", ";
  refEl.value += b + " "; // 사용자가 뒤에 3:16 등 붙이기 쉽게 스페이스만
  bookEl.value = "";
});

goBtn.addEventListener("click", async () => {
  try {
    const items = splitToItems(refEl.value);
    if (!items.length) { alert("성경 범위를 입력하세요."); return; }

    // 옵션
    const blurPx   = parseInt(blurPxEl.value || "0", 10);
    const maxChars = parseInt(maxCharsEl.value || "20", 10);
    const baseFont = parseInt(baseFontEl.value || "32", 10);
    const bodyColor  = bodyColorEl.value.replace("#","").toUpperCase();
    const titleColor = titleColorEl.value.replace("#","").toUpperCase();
    const showVerseNo     = !!showVerseNoEl.checked;
    const oneVersePerSlide= !!oneVersePerSlideEl.checked;
    const makeTitleSlide  = !!makeTitleSlideEl.checked;

    // 배경
    let bgDataURL;
    if (bgFileEl.files && bgFileEl.files[0]) {
      bgDataURL = await imgToDataURLWithBlur(bgFileEl.files[0], blurPx);
    } else {
      bgDataURL = await imgToDataURLWithBlur("BG_image.jpg", blurPx);
    }

    // 파싱 & 구절 로드
    const blocks = [];
    for (const raw of items) {
      const ref = parseRef(raw);
      const verses = await fetchVerses(ref.book, ref.sCh, ref.sV, ref.eCh, ref.eV);
      if (!verses.length) {
        throw new Error(`구절을 찾지 못했습니다: ${ref.book} ${ref.sCh}:${ref.sV}~${ref.eCh}:${ref.eV}`);
      }
      blocks.push({ ref, verses });
    }

    // PPT 생성
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "LAYOUT_16x9", width: 13.33, height: 7.5 });
    pptx.layout = "LAYOUT_16x9";

    if (makeTitleSlide) {
      const tSlide = pptx.addSlide();
      tSlide.background = { data: bgDataURL };
      const title = items.join(" / ");
      tSlide.addText(title, {
        x:0, y:2.8, w:"100%", h:1.5, fontSize:44, bold:true, align:"center",
        color: titleColor, fontFace: "Malgun Gothic"
      });
    }

    for (const blk of blocks) {
      const { ref, verses } = blk;
      const titleText = `${ref.book} ${ref.sCh}:${ref.sV}` + (ref.eCh===ref.sCh && ref.eV===ref.sV ? "" : `~${ref.eCh===ref.sCh ? ref.eV : ref.eCh + ":" + ref.eV}`);

      if (oneVersePerSlide) {
        for (const vs of verses) {
          const s = pptx.addSlide();
          s.background = { data: bgDataURL };
          s.addText(titleText, {
            x:0, y:0.3, w:"100%", h:0.8, fontSize:40, bold:true, align:"center",
            color: titleColor, fontFace: "Malgun Gothic"
          });

          const verseLabel = showVerseNo ? `${vs.ch}:${vs.v} ` : "";
          const content = verseLabel + vs.text;
          const lines = linesForVerse(content, maxChars);
          const fontSize = autoFont(baseFont, content.length);

          s.addText(lines.join("\n"), {
            x:1, y:1.2, w:11.33, h:5.5,
            fontSize: fontSize, bold:true, color: bodyColor,
            align:"center", valign:"middle",
            fontFace: "Dotum"
          });
        }
      } else {
        const s = pptx.addSlide();
        s.background = { data: bgDataURL };
        s.addText(titleText, {
          x:0, y:0.3, w:"100%", h:0.8, fontSize:40, bold:true, align:"center",
          color: titleColor, fontFace: "Malgun Gothic"
        });

        const merged = verses.map(v => (showVerseNo ? `${v.ch}:${v.v} ` : "") + v.text).join(" ");
        const lines = linesForVerse(merged, maxChars);
        const fontSize = autoFont(baseFont, merged.length);

        s.addText(lines.join("\n"), {
          x:1, y:1.2, w:11.33, h:5.5,
          fontSize: fontSize, bold:true, color: bodyColor,
          align:"center", valign:"middle",
          fontFace: "Dotum"
        });
      }
    }

    const filename = `성경슬라이드_${Date.now()}.pptx`;
    await pptx.writeFile({ fileName: filename });
  } catch (err) {
    console.error(err);
    alert(err.message || String(err));
  }
});
