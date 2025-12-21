<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>CT 통계 (사진 순서 그대로)</title>
  <style>
    body { font-family: Arial, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif; margin: 14px; color:#111; }
    .controls { border:1px solid #ddd; border-radius:12px; padding:12px; margin-bottom:12px; }
    .row { display:flex; flex-wrap:wrap; gap:10px; align-items:end; }
    input[type="file"]{ display:none; }
    input[type="text"], input[type="time"]{ border:1px solid #ddd; border-radius:10px; padding:8px 10px; }
    button{ border:1px solid #ddd; border-radius:10px; padding:10px 12px; background:#fff; cursor:pointer; }
    button:active{ transform: translateY(1px); }
    label.chk { display:flex; align-items:center; gap:6px; font-size:12px; user-select:none; }

    .drop { border:2px dashed #cfcfcf; border-radius:12px; padding:10px 12px; background:#fff; min-width: 260px; }
    .drop strong{ display:block; font-size:13px; margin-bottom:4px; }
    .drop .hint{ font-size:12px; color:#666; }
    .drop.dragover{ border-color:#111; background:#fafafa; }
    .drop .file{ margin-top:6px; font-size:12px; color:#111; }

    .report { border:1px solid #000; padding:12px; }
    .header { display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:8px; }
    .title { font-size:20px; font-weight:700; text-align:center; flex:1; }
    .date { font-size:13px; font-weight:700; }

    table { width:100%; border-collapse:collapse; font-size:12px; }
    th, td { border:1px solid #000; padding:5px 6px; text-align:center; }
    th { background:#f5f5f5; font-weight:700; }
    .left { text-align:left; }
    .section { background:#f0f0f0; font-weight:800; }
    .sum { background:#fafafa; font-weight:800; }
    .muted { color:#666; font-size:12px; }
    .warn { color:#b00020; font-weight:800; }
    .ok { color:#0a7a2f; font-weight:800; }

    @media print {
      .controls { display:none; }
      body { margin: 0; }
      .report { border:none; padding:0; }
    }
  </style>
</head>
<body>
  <div class="controls">
    <div class="row">
      <div class="drop" id="drop-std">
        <strong>1) 기준정보 파일(부위)</strong>
        <div class="hint">부위 / 1~5번 / 건수 (사진 순서 그대로)</div>
        <div class="file" id="name-std">선택된 파일 없음</div>
        <input id="fStd" type="file" accept=".xlsx,.xls" />
      </div>

      <div class="drop" id="drop-ledger">
        <strong>2) 검사대장 파일</strong>
        <div class="hint">R, 시각, 검사명, 횟수, 검사자1...</div>
        <div class="file" id="name-ledger">선택된 파일 없음</div>
        <input id="fLedger" type="file" accept=".xlsx,.xls" />
      </div>

      <div class="drop" id="drop-contrast">
        <strong>3) 조영제통계(선택)</strong>
        <div class="hint">조영제 구획 숫자 채우기용</div>
        <div class="file" id="name-contrast">선택된 파일 없음</div>
        <input id="fContrast" type="file" accept=".xlsx,.xls" />
      </div>

      <div style="min-width:300px;">
        <div class="muted">야간 근무자 이름(최대 4명, 콤마)</div>
        <input id="nightNames" type="text" placeholder="예: 홍길동, 김철수, 박영희, 최민수" />
        <div class="muted">※ 주간/야간 구획 채우기용(야간 이름 우선)</div>
      </div>

      <div style="min-width:220px;">
        <div class="muted">보고서 날짜(표시용)</div>
        <input id="reportDate" type="text" placeholder="예: 2025년 12월 21일" />
      </div>

      <button id="run">사진 순서로 생성</button>
      <button id="printBtn">인쇄 / PDF</button>
    </div>

    <div id="status" class="muted" style="margin-top:8px;"></div>
  </div>

  <div class="report">
    <div class="header">
      <div style="width:120px;"></div>
      <div class="title">CT실 일일 검사 통계</div>
      <div class="date" id="dateText"></div>
    </div>

    <table>
      <thead>
        <tr>
          <th>부위</th>
          <th>1번</th><th>2번</th><th>3번</th><th>4번</th><th>5번</th>
          <th>건수</th>
        </tr>
      </thead>
      <tbody id="outBody"></tbody>
    </table>

    <div id="note" class="muted" style="margin-top:10px;"></div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script>
    const $ = (id) => document.getElementById(id);
    const norm = (v) => (v == null) ? "" : String(v).trim();

    // 기준정보 파일 컬럼명(네가 올린 파일 그대로)
    const STD = { part:"부위", r1:"1번", r2:"2번", r3:"3번", r4:"4번", r5:"5번", total:"건수" };

    // 검사대장 컬럼명(네가 말한 그대로)
    const LEDGER = { room:"R", exam:"검사명", cnt:"횟수", tech:"검사자1" };

    // 조영제 통계 컬럼(네가 준 구조)
    const CONTRAST = {
      perfName:"행위수행부서명",
      cName:"조영제코드명",
      dose:"투여량",
      unit:"투여단위"
    };

    function bindDrop(dropId, inputId, nameId){
      const drop = $(dropId);
      const input = $(inputId);
      const name = $(nameId);
      const setName = () => name.textContent = input.files?.[0]?.name || "선택된 파일 없음";

      drop.addEventListener("click", () => input.click());
      input.addEventListener("change", setName);
      drop.addEventListener("dragover", (e)=>{ e.preventDefault(); drop.classList.add("dragover"); });
      drop.addEventListener("dragleave", ()=> drop.classList.remove("dragover"));
      drop.addEventListener("drop", (e)=>{
        e.preventDefault(); drop.classList.remove("dragover");
        const files = e.dataTransfer?.files;
        if(!files || !files.length) return;
        const dt = new DataTransfer();
        dt.items.add(files[0]);
        input.files = dt.files;
        setName();
      });
      setName();
    }

    bindDrop("drop-std","fStd","name-std");
    bindDrop("drop-ledger","fLedger","name-ledger");
    bindDrop("drop-contrast","fContrast","name-contrast");

    async function readFirstSheet(file){
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(ws, { defval:"" });
    }

    function toNum(v){
      const n = Number(String(v ?? "").replace(/,/g,""));
      return Number.isFinite(n) ? n : 0;
    }

    // R 파싱: "41"->4, "31"->3, "R3"->3, "4"->4
    function extractRoomNo(rValue){
      const s = String(rValue ?? "").trim();
      if(!s) return null;
      const nums = s.match(/\d+/g);
      if(!nums) return null;
      const n = Number(nums[0]);
      if(Number.isFinite(n)){
        if(n >= 10){
          const room = Math.floor(n/10);
          if(room >= 1 && room <= 5) return room;
        }
        if(n >= 1 && n <= 5) return n;
      }
      const m = s.match(/[1-5]/);
      return m ? Number(m[0]) : null;
    }

    function parseNameSet(text){
      return new Set(String(text||"").split(",").map(x=>x.trim()).filter(Boolean));
    }

    // 기준정보 행이 "구획/헤더"인지 판별 (조영제, 주간/야간 같이 숫자 대신 '-'가 있는 행)
    function isSectionRow(stdRow){
      const part = norm(stdRow[STD.part]);
      const cells = [STD.r1,STD.r2,STD.r3,STD.r4,STD.r5,STD.total].map(k=>norm(stdRow[k]));
      const allDashOrEmpty = cells.every(v => v === "" || v === "-" );
      // 예: "조영제", "주간/야간" 같은 행은 숫자가 비어있고 구획명만 있음
      return part && allDashOrEmpty;
    }

    function isSumRow(stdRow){
      return norm(stdRow[STD.part]) === "합계";
    }

    // 검사대장 검사명 -> 기준정보 부위 매칭 (기본: 포함 매칭)
    function matchPart(partName, examName){
      if(!partName) return false;
      if(!examName) return false;
      return examName.includes(partName); // 필요하면 나중에 “정확일치” 옵션으로 바꿀 수 있음
    }

    $("printBtn").addEventListener("click", () => window.print());

    $("run").addEventListener("click", async () => {
      const fStd = $("fStd").files?.[0];
      const fLedger = $("fLedger").files?.[0];
      const fContrast = $("fContrast").files?.[0] || null;

      if(!fStd || !fLedger){
        $("status").innerHTML = `<span class="warn">기준정보 파일 + 검사대장 파일은 꼭 업로드해야 해.</span>`;
        return;
      }

      $("status").textContent = "읽는 중…";

      const stdRows = await readFirstSheet(fStd);
      const ledgerRows = await readFirstSheet(fLedger);
      const nightSet = parseNameSet($("nightNames").value);

      // ---- 1) 검사대장으로: "부위(기준정보 행)"별 방 건수 계산 ----
      // partAgg: part -> {1..5,total}
      const partAgg = new Map();
      const partsInStd = stdRows.map(r => norm(r[STD.part])).filter(Boolean);

      // initialize only for "실제 부위행"(구획/합계 제외)
      for(const r of stdRows){
        const p = norm(r[STD.part]);
        if(!p) continue;
        if(isSectionRow(r) || isSumRow(r)) continue;
        partAgg.set(p, {1:0,2:0,3:0,4:0,5:0,total:0});
      }

      // 주/야(근무자 기준) 집계 (기준정보의 "주간/야간" 섹션에 채울 용도)
      const dayAgg = {1:0,2:0,3:0,4:0,5:0,total:0};
      const nightAgg = {1:0,2:0,3:0,4:0,5:0,total:0};
      const allAgg = {1:0,2:0,3:0,4:0,5:0,total:0};

      for(const row of ledgerRows){
        const room = extractRoomNo(row[LEDGER.room]);
        if(!room) continue;
        const cnt = Math.max(1, toNum(row[LEDGER.cnt]) || 1);
        const exam = norm(row[LEDGER.exam]);
        const tech = norm(row[LEDGER.tech]);

        // 전체/주야 합산
        allAgg[room] += cnt; allAgg.total += cnt;

        const isNight = (nightSet.size > 0 && tech && nightSet.has(tech));
        if(isNight){ nightAgg[room] += cnt; nightAgg.total += cnt; }
        else { dayAgg[room] += cnt; dayAgg.total += cnt; }

        // 기준정보 부위행과 매칭되는 것만 누적
        // (포함 매칭)
        for(const [part, agg] of partAgg.entries()){
          if(matchPart(part, exam)){
            agg[room] += cnt;
            agg.total += cnt;
            break; // 한 검사명은 한 부위행에만 들어간다고 가정(원하면 다중매칭도 가능)
          }
        }
      }

      // ---- 2) 조영제통계로: "조영제 섹션" 행 채우기(선택) ----
      // std의 조영제 행들은 보통 "Bonorex 350 inj [100ml]" 같은 이름이 part 컬럼에 있음
      const contrastAgg = new Map(); // label -> count(total only)
      if(fContrast){
        const contrastRows = await readFirstSheet(fContrast);
        for(const r of contrastRows){
          // CT만 남기고 싶으면 perfName 필터를 여기에 추가 가능
          const cName = norm(r[CONTRAST.cName]);
          if(!cName) continue;
          const dose = norm(r[CONTRAST.dose]);
          const unit = norm(r[CONTRAST.unit]);
          const label = `${cName}${dose||unit ? ` inj [${dose}${unit?unit:""}]` : ""}`.replace(/\s+/g," ").trim();
          contrastAgg.set(label, (contrastAgg.get(label)||0) + 1);
        }
      }

      // ---- 3) 기준정보 순서대로 "사진처럼" 출력 ----
      const tbody = $("outBody");
      tbody.innerHTML = "";

      let currentSection = "부위"; // "부위" | "조영제" | "주야" 등
      // 섹션별 소계용 러닝합
      let run = {1:0,2:0,3:0,4:0,5:0,total:0};

      const flushSumRow = () => {
        const tr = document.createElement("tr");
        tr.className = "sum";
        tr.innerHTML = `
          <td class="left">합계</td>
          <td>${run[1]||0}</td><td>${run[2]||0}</td><td>${run[3]||0}</td><td>${run[4]||0}</td><td>${run[5]||0}</td>
          <td>${run.total||0}</td>
        `;
        tbody.appendChild(tr);
      };

      const resetRun = () => run = {1:0,2:0,3:0,4:0,5:0,total:0};

      for(const r of stdRows){
        const label = norm(r[STD.part]);
        if(!label) continue;

        // 섹션 헤더(예: 조영제, 주간/야간)
        if(isSectionRow(r)){
          // 섹션 전환 전에는 러닝 합계를 리셋(기존 섹션 합계는 기준정보에 '합계' 행이 있으므로 거기서 출력됨)
          currentSection = label.includes("조영") ? "조영제" : (label.includes("주간") || label.includes("야간") || label.includes("주/야") ? "주야" : label);
          const tr = document.createElement("tr");
          tr.className = "section";
          tr.innerHTML = `
            <td class="left">${label}</td>
            <td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td>
          `;
          tbody.appendChild(tr);
          resetRun();
          continue;
        }

        // 합계행: 지금까지 run을 출력
        if(isSumRow(r)){
          flushSumRow();
          resetRun();
          continue;
        }

        // 일반 데이터 행: 섹션에 따라 숫자 채우기
        let v = {1:0,2:0,3:0,4:0,5:0,total:0};

        if(currentSection === "조영제"){
          // 조영제는 방별이 아니라 총계만(사진처럼 방칸은 0 또는 공란도 가능)
          const cnt = contrastAgg.get(label) || 0;
          v.total = cnt;
          // 사진처럼 방칸은 0으로 두고 싶으면 아래를 그대로, 공란이면 ""로 바꿔도 됨
          v[1]=0; v[2]=0; v[3]=0; v[4]=0; v[5]=0;
        } else if(currentSection.includes("주") && currentSection.includes("야") || currentSection === "주야") {
          // "주간", "야간" 같은 행이 기준정보에 실제로 들어있는 구조(네 스샷처럼)
          // 기준정보의 label이 "주간"이면 dayAgg, "야간"이면 nightAgg로 채움
          if(label === "주간"){
            v = {...dayAgg};
          } else if(label === "야간"){
            v = {...nightAgg};
          } else {
            // 혹시 다른 텍스트면 전체로
            v = {...allAgg};
          }
        } else {
          // 부위 섹션: 검사대장 기반 partAgg에서 채움
          const agg = partAgg.get(label);
          if(agg){
            v = {...agg};
          } else {
            // 기준정보에 있는데 매칭이 안 된 항목은 0으로 표시
            v = {1:0,2:0,3:0,4:0,5:0,total:0};
          }
        }

        // 러닝 합계 누적(각 섹션의 합계 행을 만들어주기 위함)
        run[1]+=v[1]; run[2]+=v[2]; run[3]+=v[3]; run[4]+=v[4]; run[5]+=v[5]; run.total+=v.total;

        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td class="left">${label}</td>
          <td>${v[1] || 0}</td>
          <td>${v[2] || 0}</td>
          <td>${v[3] || 0}</td>
          <td>${v[4] || 0}</td>
          <td>${v[5] || 0}</td>
          <td>${v.total || 0}</td>
        `;
        tbody.appendChild(tr);
      }

      $("dateText").textContent = norm($("reportDate").value);
      $("status").innerHTML = `<span class="ok">완료 ✅</span> (기준정보 파일 순서 그대로 표시)`;

      $("note").innerHTML =
        `- 부위 섹션은 <b>검사대장(R, 검사명, 횟수)</b>로 채움<br/>
         - 조영제 섹션은 <b>조영제통계(선택)</b>로 총계만 채움<br/>
         - 주간/야간 섹션은 <b>검사자1 이름(야간 목록 매칭)</b>으로 채움`;
    });
  </script>
</body>
</html>
