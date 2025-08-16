// ========= 성경 권 이름 매핑 (정식명 + 약칭) =========
const BOOKS = [
  { canon:"창세기", aliases:["창"] }, { canon:"출애굽기", aliases:["출"] },
  { canon:"레위기", aliases:["레"] }, { canon:"민수기", aliases:["민"] },
  { canon:"신명기", aliases:["신"] }, { canon:"여호수아", aliases:["수"] },
  { canon:"사사기", aliases:["삿"] }, { canon:"룻기", aliases:["룻"] },
  { canon:"사무엘상", aliases:["삼상"] }, { canon:"사무엘하", aliases:["삼하"] },
  { canon:"열왕기상", aliases:["왕상"] }, { canon:"열왕기하", aliases:["왕하"] },
  { canon:"역대상", aliases:["대상"] }, { canon:"역대하", aliases:["대하"] },
  { canon:"에스라", aliases:["스"] }, { canon:"느헤미야", aliases:["느"] },
  { canon:"에스더", aliases:["에"] }, { canon:"욥기", aliases:["욥"] },
  { canon:"시편", aliases:["시"] }, { canon:"잠언", aliases:["잠"] },
  { canon:"전도서", aliases:["전"] }, { canon:"아가", aliases:["아"] },
  { canon:"이사야", aliases:["사"] }, { canon:"예레미야", aliases:["렘"] },
  { canon:"예레미야애가", aliases:["애"] }, { canon:"에스겔", aliases:["겔"] },
  { canon:"다니엘", aliases:["단"] }, { canon:"호세아", aliases:["호"] },
  { canon:"요엘", aliases:["욜"] }, { canon:"아모스", aliases:["암"] },
  { canon:"오바댜", aliases:["옵"] }, { canon:"요나", aliases:["욘"] },
  { canon:"미가", aliases:["미"] }, { canon:"나훔", aliases:["나"] },
  { canon:"하박국", aliases:["합"] }, { canon:"스바냐", aliases:["습"] },
  { canon:"학개", aliases:["학"] }, { canon:"스가랴", aliases:["슥"] },
  { canon:"말라기", aliases:["말"] },

  { canon:"마태복음", aliases:["마"] }, { canon:"마가복음", aliases:["막"] },
  { canon:"누가복음", aliases:["눅"] }, { canon:"요한복음", aliases:["요"] },
  { canon:"사도행전", aliases:["행"] }, { canon:"로마서", aliases:["롬"] },
  { canon:"고린도전서", aliases:["고전"] }, { canon:"고린도후서", aliases:["고후"] },
  { canon:"갈라디아서", aliases:["갈"] }, { canon:"에베소서", aliases:["엡"] },
  { canon:"빌립보서", aliases:["빌"] }, { canon:"골로새서", aliases:["골"] },
  { canon:"데살로니가전서", aliases:["살전"] }, { canon:"데살로니가후서", aliases:["살후"] },
  { canon:"디모데전서", aliases:["딤전"] }, { canon:"디모데후서", aliases:["딤후"] },
  { canon:"디도서", aliases:["딛"] }, { canon:"빌레몬서", aliases:["몬"] },
  { canon:"히브리서", aliases:["히"] }, { canon:"야고보", aliases:["약"] },
  { canon:"베드로전서", aliases:["벧전"] }, { canon:"베드로후서", aliases:["벧후"] },
  { canon:"요한일서", aliases:["요일"] }, { canon:"요한이서", aliases:["요이"] },
  { canon:"요한삼서", aliases:["요삼"] }, { canon:"유다서", aliases:["유"] },
  { canon:"요한계시록", aliases:["계"] },
];
const BOOK_ALIAS_TO_CANON = (() => {
  const map = new Map();
  BOOKS.forEach(({ canon, aliases }) => {
    map.set(canon, canon);
    aliases.forEach(a => map.set(a, canon));
  });
  return map;
})();

// ========= 고정 스타일/옵션 =========
const STYLE = {
  titleColor: "558ED5",
  bodyColor:  "FFFF00",
  maxChars:   20,
  baseFont:   32,
  fontFaceBody:  "Dotum",
  fontFaceTitle: "Malgun Gothic",
  showVerseNo: true,
  outputName:  "BibleVerses.pptx",
  background:  "BG_image.jpg", // 있으면 사용
};

// ========= DOM =========
const refEl    = document.getElementById("reference");
const statusEl = document.getElementById("status");

// ========= 로그/진행 =========
function clearStatus(){ statusEl.innerHTML = ""; statusEl.style.display = 'block'; }
function logStatus(msg){
  const p = document.createElement("p");
  p.textContent = msg;
  statusEl.appendChild(p);
  statusEl.scrollTop = statusEl.scrollHeight;
}
function setProgress(cur, total){
  const pct = total ? Math.round(cur*100/total) : 0;
  logStatus(`진행률: ${pct}% (${cur}/${total})`);
}

// ========= 유틸 =========
function trimAll(s){ return (s||"").replace(/\s+/g," ").trim(); }
function splitItems(input){
  return input.split(/[\n,]+/).map(s=>trimAll(s)).filter(Boolean);
}
function normalizeBook(name){
  const canon = BOOK_ALIAS_TO_CANON.get(trimAll(name));
  if (!canon) throw new Error(`알 수 없는 성경 이름: "${name}"`);
  return canon;
}
// A) 요 3:16~4:3  B) 요 3:16~18  C) 요 3:16
function parseRef(text){
  const t = trimAll(text);
  const pA = /^(.+?)\s+(\d+):(\d+)\s*~\s*(\d+):(\d+)$/;
  const pB = /^(.+?)\s+(\d+):(\d+)\s*~\s*(\d+)$/;
  const pC = /^(.+?)\s+(\d+):(\d+)$/;

  let m = t.match(pA);
  if (m) return {book:normalizeBook(m[1]), sCh:+m[2], sV:+m[3], eCh:+m[4], eV:+m[5]};
  m = t.match(pB);
  if (m) return {book:normalizeBook(m[1]), sCh:+m[2], sV:+m[3], eCh:+m[2], eV:+m[4]};
  m = t.match(pC);
  if (m) return {book:normalizeBook(m[1]), sCh:+m[2], sV:+m[3], eCh:+m[2], eV:+m[3]};
  throw new Error(`입력 형식 오류: "${text}" (예: 요 3:16, 요 3:16~18, 요 3:16~4:3)`);
}

async function exists(url){
  try {
    const res = await fetch(encodeURI(url), { method: "HEAD" });
    return res.ok;
  } catch { return false; }
}

// 텍스트 정리: 줄바꿈 통일 + XML에 허용되지 않는 제어문자 제거
function sanitizeText(s){
  return String(s)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")  // XML 금지 제어문자 제거
    .trim();
}

async function fetchVerses(book, sCh, sV, eCh, eV){
  const path = `bible/${book}.txt`;
  const res = await fetch(encodeURI(path));
  if (!res.ok) throw new Error(`성경 파일을 찾을 수 없습니다: ${path} (HTTP ${res.status})`);
  const txt = await res.text();
  const lines = txt.split(/\r?\n/);
  const out = [];
  for (const raw of lines){
    const line = sanitizeText(raw);
    const m = line.match(/^(\d+):(\d+)\s+(.*)$/);
    if (!m) continue;
    const ch = +m[1], v = +m[2], body = m[3].trim();
    if ((ch > eCh) || (ch === eCh && v > eV)) break;
    if ((ch > sCh) || (ch === sCh && v >= sV)) out.push({ch, v, text: body});
  }
  return out;
}

function chunkBy(s, maxLen){
  const arr = [];
  for (let i=0;i<s.length;i+=maxLen) arr.push(s.slice(i, i+maxLen));
  return arr;
}
function autoFont(base, total){
  if (total <= 40) return base;
  const dec = Math.min(14, Math.floor((total-40)*0.3));
  return Math.max(18, base - dec);
}

// 배경 이미지를 data URL로 로드(안전)
async function loadBgAsDataURL(path){
  const res = await fetch(encodeURI(path));
  if (!res.ok) throw new Error(`배경 이미지를 불러올 수 없습니다: ${path} (HTTP ${res.status})`);
  const blob = await res.blob();
  return await new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsDataURL(blob);
  });
}

// ========= 메인: 버튼에서 호출 =========
window.generatePPT = async function generatePPT(){
  clearStatus();
  logStatus("생성 시작");

  try{
    if (typeof window.PptxGenJS === "undefined"){
      throw new Error("PptxGenJS가 로드되지 않았습니다. (CDN/로컬 경로 확인)");
    }
    const raw = sanitizeText(refEl.value || "");
    if (!raw) throw new Error("성경 구절을 입력하세요.");

    // 1) 파싱
    logStatus("입력값 파싱 중…");
    const refs = splitItems(raw).map(parseRef);
    logStatus(`입력 구간 수: ${refs.length}`);

    // 2) 배경 준비 (data URL)
    let bgDataUrl = null;
    if (await exists(STYLE.background)) {
      logStatus(`배경 감지: ${STYLE.background}`);
      bgDataUrl = await loadBgAsDataURL(STYLE.background);
    }

    // 3) 본문 로드
    logStatus("성경 본문 로드…");
    const blocks = [];
    for (const ref of refs){
      logStatus(`- 로드: ${ref.book} ${ref.sCh}:${ref.sV}~${ref.eCh}:${ref.eV}`);
      const verses = await fetchVerses(ref.book, ref.sCh, ref.sV, ref.eCh, ref.eV);
      if (!verses.length) throw new Error(`구절을 찾지 못했습니다: ${ref.book} ${ref.sCh}:${ref.sV}~${ref.eCh}:${ref.eV}`);
      blocks.push({ref, verses});
    }
    logStatus("본문 로드 완료");

    // 4) PPT 생성
    const totalSlides = blocks.reduce((sum,b)=> sum + b.verses.length, 0);
    let made = 0;
    setProgress(made, totalSlides);
    logStatus("PPT 슬라이드 생성 시작…");

    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "LAYOUT_16x9", width: 13.33, height: 7.5 });
    pptx.layout = "LAYOUT_16x9";

    for (const blk of blocks){
      const { ref, verses } = blk;
      const title =
        `${ref.book} ${ref.sCh}:${ref.sV}` +
        (ref.eCh===ref.sCh && ref.eV===ref.sV ? "" : `~${ref.eCh===ref.sCh ? ref.eV : ref.eCh+":"+ref.eV}`);

      for (const vs of verses){
        const s = pptx.addSlide();

        if (bgDataUrl) s.background = { data: bgDataUrl };

        // 제목
        s.addText(title, {
          x:0, y:0.3, w:"100%", h:0.8,
          fontSize:40, bold:true, align:"center",
          color:STYLE.titleColor, fontFace:STYLE.fontFaceTitle
        });

        // 본문
        const head = STYLE.showVerseNo ? `${vs.ch}:${vs.v} ` : "";
        const content = sanitizeText(head + vs.text);
        const lines = chunkBy(content, STYLE.maxChars);
        const fz = autoFont(STYLE.baseFont, content.length);

        s.addText(lines.join("\n"), {
          x:1, y:1.2, w:11.33, h:5.5,
          fontSize:fz, bold:true, color:STYLE.bodyColor,
          align:"center", valign:"middle", fontFace:STYLE.fontFaceBody
        });

        made++;
        if (made % 5 === 0 || made === totalSlides) setProgress(made, totalSlides);
      }
    }

    logStatus("슬라이드 생성 완료");
    logStatus("파일 저장…");

    // ✅ 버전 호환 안전: 객체 인자로 파일명 지정
    await pptx.writeFile({ fileName: STYLE.outputName });

    logStatus(`✅ 완료: ${STYLE.outputName}`);

  }catch(err){
    console.error(err);
    logStatus(`❌ 오류: ${err.message || String(err)}`);
    alert(err.message || String(err));
  }
};
