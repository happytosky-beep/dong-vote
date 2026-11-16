/**
 * 이문아이파크자이 3단지 입대의 액션리스트 v2 — 구글 시트 저장소 (Apps Script)
 * 액션리스트(items) + 입주민 의견함(opinions)
 *
 * [처음 설치]
 * 1) 구글 시트를 새로 만들고 [확장 프로그램 → Apps Script]에 이 코드를 통째로 붙여넣기
 * 2) 아래 PIN_INIT 을 입대의 편집 비밀번호로 바꾸기
 * 3) 함수 선택에서 setup 실행 (권한 허용) → items 시트에 31개 항목, opinions 시트 생성
 * 4) [배포 → 새 배포 → 웹 앱] 실행 사용자: 나 / 액세스: 모든 사용자 → 웹 앱 URL 복사
 *
 * [v1에서 업그레이드]
 * 코드를 이것으로 바꾸고 setup 한 번 실행(기존 항목은 그대로 둠) →
 * [배포 → 배포 관리 → 수정(연필) → 버전: 새 버전 → 배포] (URL 그대로 유지)
 */
const PIN_INIT = '여기에_비밀번호';
const ITEMS = 'items', OPS = 'opinions';
const COLS = ['id','no','cat','title','contacts','quick','resolve','owner','proposer','priority','status','letter','agenda','due','note','updatedAt'];
const OCOLS = ['id','no','cat','title','body','dong','relNo','status','reply','linkedNo','likes','likedBy','createdAt','updatedAt'];

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 2000, cols.length).setNumberFormat('@');
    s.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v) { return v === undefined || v === null ? '' : String(v); }

function toRow_(o) {
  return COLS.map(k => {
    let v = o[k];
    if (k === 'contacts') v = Array.isArray(v) ? v.join(', ') : (v || '');
    if (k === 'letter' || k === 'agenda') v = v ? 'Y' : '';
    return str_(v);
  });
}
function fromRow_(r) {
  const o = {}; COLS.forEach((k, i) => o[k] = r[i]);
  o.no = Number(o.no) || 0;
  o.contacts = String(o.contacts || '').split(',').map(x => x.trim()).filter(Boolean);
  o.letter = o.letter === 'Y'; o.agenda = o.agenda === 'Y';
  o._id = String(o.id);
  return o;
}
function toORow_(o) { return OCOLS.map(k => str_(o[k])); }
function fromORow_(r) {
  const o = {}; OCOLS.forEach((k, i) => o[k] = r[i]);
  o.no = Number(o.no) || 0; o.likes = Number(o.likes) || 0;
  delete o.likedBy; o._id = String(o.id);
  return o;
}
function rows_(s) { const v = s.getDataRange().getValues(); return v.length < 2 ? [] : v.slice(1).filter(r => r[0]); }

function setup() {
  PropertiesService.getScriptProperties().setProperty('PIN', PIN_INIT);
  const s = sheet_(ITEMS, COLS);
  if (s.getLastRow() < 2) {
    const now = new Date().toISOString();
    const rows = SEED.map(o => toRow_(Object.assign({ updatedAt: now }, o)));
    s.getRange(2, 1, rows.length, COLS.length).setValues(rows);
  }
  sheet_(OPS, OCOLS);
}

function doGet() {
  return json_({
    ok: true,
    items: rows_(sheet_(ITEMS, COLS)).map(fromRow_),
    opinions: rows_(sheet_(OPS, OCOLS)).map(fromORow_),
    at: new Date().toISOString()
  });
}

function findRow_(s, id) {
  const v = s.getDataRange().getValues();
  const i = v.findIndex((r, n) => n > 0 && String(r[0]) === String(id));
  return { v: v, i: i };
}
function maxNo_(v) { return v.slice(1).reduce((m, r) => Math.max(m, Number(r[1]) || 0), 0); }

function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: '잘못된 요청' }); }
  const now = new Date().toISOString();
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    // ---- 입주민 (비밀번호 없음) ----
    if (b.action === 'opinion_add') {
      const it = b.item || {};
      const title = str_(it.title).trim().slice(0, 80), body = str_(it.body).trim().slice(0, 1500);
      if (!title || !body) return json_({ ok: false, error: '제목과 내용을 입력하세요' });
      const cid = str_(b.cid).slice(0, 40);
      const cache = CacheService.getScriptCache();
      if (cid && cache.get('op_' + cid)) return json_({ ok: false, error: '잠시 후 다시 등록해 주세요 (1분 간격)' });
      const s = sheet_(OPS, OCOLS);
      const o = { id: 'o' + Date.now(), no: maxNo_(s.getDataRange().getValues()) + 1, cat: str_(it.cat).slice(0, 30),
        title: title, body: body, dong: str_(it.dong).slice(0, 10), relNo: str_(it.relNo).slice(0, 10),
        status: '접수', reply: '', linkedNo: '', likes: 0, likedBy: '', createdAt: now, updatedAt: now };
      s.appendRow(toORow_(o));
      if (cid) cache.put('op_' + cid, '1', 60);
      return json_({ ok: true });
    }
    if (b.action === 'opinion_like') {
      const cid = str_(b.cid).slice(0, 40); if (!cid) return json_({ ok: false, error: '잘못된 요청' });
      const s = sheet_(OPS, OCOLS); const f = findRow_(s, b.id);
      if (f.i < 1) return json_({ ok: false, error: '의견을 찾지 못했습니다' });
      const r = f.v[f.i]; const li = OCOLS.indexOf('likedBy'), lk = OCOLS.indexOf('likes');
      const by = String(r[li] || '').split(',').filter(Boolean);
      const has = by.indexOf(cid) >= 0;
      const next = has ? by.filter(x => x !== cid) : by.concat([cid]);
      s.getRange(f.i + 1, lk + 1).setValue(String(next.length));
      s.getRange(f.i + 1, li + 1).setValue(next.join(','));
      return json_({ ok: true, likes: next.length, liked: !has });
    }

    // ---- 입대의 (편집 비밀번호 필요) ----
    const pin = PropertiesService.getScriptProperties().getProperty('PIN');
    if (!pin || String(b.pin || '') !== pin) return json_({ ok: false, error: '편집 비밀번호가 맞지 않습니다' });

    const si = sheet_(ITEMS, COLS);
    if (b.action === 'check') return json_({ ok: true });
    if (b.action === 'add') {
      const it = Object.assign({}, b.item, { id: 'i' + Date.now(), no: maxNo_(si.getDataRange().getValues()) + 1, updatedAt: now });
      si.appendRow(toRow_(it));
      return json_({ ok: true, item: it });
    }
    if (b.action === 'update' || b.action === 'delete') {
      const f = findRow_(si, b.id); if (f.i < 1) return json_({ ok: false, error: '항목을 찾지 못했습니다' });
      if (b.action === 'delete') { si.deleteRow(f.i + 1); return json_({ ok: true }); }
      const cur = fromRow_(f.v[f.i]);
      const it = Object.assign(cur, b.item, { id: cur.id, no: cur.no, updatedAt: now });
      si.getRange(f.i + 1, 1, 1, COLS.length).setValues([toRow_(it)]);
      return json_({ ok: true });
    }

    const so = sheet_(OPS, OCOLS);
    const f = findRow_(so, b.id); if (f.i < 1) return json_({ ok: false, error: '의견을 찾지 못했습니다' });
    const raw = f.v[f.i]; const cur = {}; OCOLS.forEach((k, i) => cur[k] = raw[i]);
    if (b.action === 'opinion_delete') { so.deleteRow(f.i + 1); return json_({ ok: true }); }
    if (b.action === 'opinion_update') {
      const p = b.item || {};
      ['status', 'reply', 'linkedNo'].forEach(k => { if (p[k] !== undefined) cur[k] = str_(p[k]); });
      cur.updatedAt = now;
      so.getRange(f.i + 1, 1, 1, OCOLS.length).setValues([toORow_(cur)]);
      return json_({ ok: true });
    }
    if (b.action === 'opinion_promote') {
      const no = maxNo_(si.getDataRange().getValues()) + 1;
      const it = { id: 'i' + Date.now(), no: no, cat: cur.cat || '기타', title: cur.title, contacts: [], quick: '', resolve: '',
        owner: '', proposer: '입주민 의견 💬' + cur.no + (cur.dong ? ' (' + cur.dong + ')' : ''), priority: '보통', status: '미착수',
        letter: false, agenda: false, due: '', note: '입주민 의견 원문: ' + cur.body, updatedAt: now };
      si.appendRow(toRow_(it));
      cur.status = '액션리스트 등록'; cur.linkedNo = String(no); cur.updatedAt = now;
      so.getRange(f.i + 1, 1, 1, OCOLS.length).setValues([toORow_(cur)]);
      return json_({ ok: true, no: no });
    }
    return json_({ ok: false, error: '알 수 없는 작업' });
  } finally { lock.releaseLock(); }
}

const SEED = [
{
"agenda": true,
"cat": "입대의 운영·계약",
"contacts": [
"선거관리위원회",
"동대문구청 공동주택과"
],
"due": "",
"letter": true,
"no": 1,
"note": "11/16(일) 19시 선관위 주재 임원 선출 회의",
"owner": "회장 / 총무이사",
"priority": "높음",
"proposer": "Jumbo",
"quick": "11/16 선관위 주재 선출 결과를 회의록에 기록하고 단톡방·게시판에 공지",
"resolve": "구청 신고 대상 여부 확인 후 입주자대표회의 구성(임원) 신고 완료",
"status": "진행중",
"title": "임원 선출(회장·감사·총무이사) 결과 정리 및 신고",
"id": "a01"
},
{
"agenda": true,
"cat": "입대의 운영·계약",
"contacts": [
"한국전력",
"관리사무소"
],
"due": "",
"letter": false,
"no": 2,
"note": "시뮬레이션 필요",
"owner": "재무이사",
"priority": "보통",
"proposer": "Jumbo",
"quick": "관리사무소에 최근 전기요금 고지서·세대별 사용량 자료 요청, 한전에 계약방식별 요금 시뮬레이션 요청",
"resolve": "종합/단일 비교표 작성 → 변경 요건(입주자 동의 등) 확인 → 입대의 의결 → 한전 변경 신청",
"status": "미착수",
"title": "전기계약방식 변경 검토 (종합 → 단일)",
"id": "a02"
},
{
"agenda": true,
"cat": "입대의 운영·계약",
"contacts": [
"현 관리업체",
"동대문구청 공동주택과"
],
"due": "",
"letter": false,
"no": 3,
"note": "",
"owner": "총무이사 / 감사",
"priority": "높음",
"proposer": "Jumbo",
"quick": "현 관리 도급계약서의 계약기간·해지조건 확인, 국토부 「주택관리업자 및 사업자 선정지침」 입찰 절차 정리",
"resolve": "관리방법 변경 의결(입주자 동의 요건 확인) → 위탁관리업체 입찰·선정 → 계약 체결",
"status": "미착수",
"title": "관리방식 변경 (도급 → 위탁) 및 위탁관리업체 선정",
"id": "a03"
},
{
"cat": "입대의 운영·계약",
"contacts": [
"위탁관리업체(선정 후)",
"현 용역업체"
],
"due": "",
"letter": false,
"no": 4,
"note": "선행: #3",
"owner": "총무이사 / 감사",
"priority": "보통",
"proposer": "Jumbo",
"quick": "현 용역별 계약 만료일·인원·금액을 표로 정리",
"resolve": "#3 위탁관리업체 선정 후 사업자 선정지침에 따라 입찰 → 입대의 의결 → 계약",
"status": "미착수",
"title": "경비(보안)·미화·조경 용역업체 선정",
"id": "a04"
},
{
"agenda": true,
"cat": "입대의 운영·계약",
"contacts": [
"조합(조합장)",
"HDC현산 CS팀"
],
"due": "",
"letter": true,
"no": 5,
"note": "",
"owner": "회장 / 총무이사",
"priority": "높음",
"proposer": "Jumbo",
"quick": "이 대시보드 '연락처별' 탭에서 조합·현산 항목을 묶어 요청 목록 1차 작성",
"resolve": "일괄 공문 발송(회신기한 명시) + 조합·현산·입대의 월 1회 정례 협의 합의",
"status": "미착수",
"title": "조합·시공사 협의사항 통합 정리 및 정기 협의",
"id": "a05"
},
{
"agenda": true,
"cat": "용역관리(청소·경비)",
"contacts": [
"관리사무소",
"청소용역업체"
],
"due": "",
"letter": false,
"no": 6,
"note": "청소업체 관리 방법도 함께 정할 것",
"owner": "총무이사",
"priority": "높음",
"proposer": "정수",
"quick": "청소 계약서·과업지시서 사본, 인원(3명으로 파악)·구역별 청소일지 요청",
"resolve": "주차장·계단이 과업에 없으면 업체와 범위 조정(인원·금액 변경 견적) → 입대의 의결. 이후 월 1회 점검표로 관리",
"status": "미착수",
"title": "청소 용역 업무범위 확인 (주차장·계단·공용부 포함 여부)",
"id": "a06"
},
{
"agenda": true,
"cat": "용역관리(청소·경비)",
"contacts": [
"관리사무소",
"경비용역업체"
],
"due": "",
"letter": false,
"no": 7,
"note": "'쉬는시간 없게'가 아니라 '휴게 중 공백 없게'로 목표 설정",
"owner": "총무이사",
"priority": "높음",
"proposer": "정수",
"quick": "경비 인원·근무형태·초소 배치·휴게시간표 요청",
"resolve": "휴게시간은 법정 보장이라 없앨 수 없음 → 교대(시차) 휴게로 초소 공백이 없도록 근무표 조정. 부족하면 인원 보강·CCTV 보완 검토",
"status": "미착수",
"title": "경비 근무체계 확인 (휴게시간 중 초소 공백 해소)",
"id": "a07"
},
{
"agenda": true,
"cat": "커뮤니티센터",
"contacts": [
"조합",
"HDC현산",
"관리사무소"
],
"due": "",
"letter": true,
"no": 8,
"note": "",
"owner": "커뮤니티이사",
"priority": "높음",
"proposer": "정수",
"quick": "관리사무소와 합동 점검 — 시설·장비·비품 목록과 하자 사진 기록",
"resolve": "하자 없는 시설부터 부분 인수 가능한지 조합·현산에 공문 질의 → 입대의 동의 범위 의결",
"status": "미착수",
"title": "커뮤니티센터 상태 점검 및 부분 인수(이관) 가능 여부",
"id": "a08"
},
{
"agenda": true,
"cat": "커뮤니티센터",
"contacts": [
"관리사무소",
"커뮤니티 위탁운영업체"
],
"due": "",
"letter": false,
"no": 9,
"note": "",
"owner": "커뮤니티이사",
"priority": "보통",
"proposer": "Jumbo",
"quick": "인근 단지 운영 사례 수집, 위탁업체 2~3곳 견적 요청",
"resolve": "비용·장단점 비교표 → 입대의 의결 → 커뮤니티 운영규정 제정",
"status": "미착수",
"title": "커뮤니티 운영방식 결정 (자치관리 vs 위탁)",
"id": "a09"
},
{
"cat": "커뮤니티센터",
"contacts": [
"조합"
],
"due": "",
"letter": true,
"no": 10,
"note": "조합장님께 부탁드릴 일",
"owner": "커뮤니티이사",
"priority": "보통",
"proposer": "장수정",
"quick": "당초 설계·장비 목록에 포함됐는지 확인",
"resolve": "조합장께 설치 요청 공문 → 설치 일정 회신",
"status": "미착수",
"title": "헬스클럽 천국의 계단 설치 요청",
"id": "a10"
},
{
"cat": "커뮤니티센터",
"contacts": [
"HDC현산 CS팀",
"관리사무소"
],
"due": "",
"letter": true,
"no": 11,
"note": "",
"owner": "커뮤니티이사 / 기술이사",
"priority": "높음",
"proposer": "장수정",
"quick": "발생 위치 사진, 환기설비 가동 상태 확인",
"resolve": "현산에 하자 접수 → 원인 조사(누수·환기) 결과 회신 → 보수 완료 확인",
"status": "미착수",
"title": "골프장·사우나 곰팡이 냄새 원인 제거",
"id": "a11"
},
{
"cat": "커뮤니티센터",
"contacts": [
"관리사무소",
"조합"
],
"due": "",
"letter": false,
"no": 12,
"note": "",
"owner": "커뮤니티이사",
"priority": "낮음",
"proposer": "장수정",
"quick": "게스트룸 현장 확인(비품·하자)",
"resolve": "이용규정·예약방식 마련 후 오픈",
"status": "미착수",
"title": "게스트룸 상태 확인 및 운영 준비",
"id": "a12"
},
{
"agenda": true,
"cat": "주차·출입",
"contacts": [
"관리사무소",
"출입통제 시스템 업체"
],
"due": "",
"letter": false,
"no": 13,
"note": "",
"owner": "기술이사",
"priority": "높음",
"proposer": "정수, 장수정",
"quick": "305동 출입 시 307동 비번이 안 되고 카드만 되는 원인 확인, 커뮤니티 출입 시 307동 입구 공용비번 사용 여부 확인",
"resolve": "출입 방식안(공용비번 vs 카드/앱) 결정 → 시스템 설정 변경 → 입주민 공지",
"status": "미착수",
"title": "출입 비밀번호·카드 체계 정리 (305·307동, 커뮤니티)",
"id": "a13"
},
{
"agenda": true,
"cat": "주차·출입",
"contacts": [
"관리사무소",
"동대문구청 공동주택과"
],
"due": "",
"letter": false,
"no": 14,
"note": "",
"owner": "재무이사",
"priority": "높음",
"proposer": "Jumbo, 장수정",
"quick": "현 관리규약 주차 조항, 세대별 등록 차량 대수 현황 받기",
"resolve": "규정안 작성 → 입대의 의결 → 관리규약 개정 사항이면 입주자 동의·구청 신고",
"status": "미착수",
"title": "주차관리규정 심의 (1가구 2대 이상 추가요금 포함)",
"id": "a14"
},
{
"cat": "주차·출입",
"contacts": [
"관리사무소"
],
"due": "",
"letter": false,
"no": 15,
"note": "",
"owner": "기술이사",
"priority": "보통",
"proposer": "장수정",
"quick": "입주민 설문(전기차 보유·구매 예정) 배포",
"resolve": "수요 기반 전기차·일반차 공용 구획안 → 입대의 의결 → 도색·표지 정비",
"status": "미착수",
"title": "전기차 수요조사 및 주차구획 정비",
"id": "a15"
},
{
"cat": "주차·출입",
"contacts": [
"조합",
"HDC현산"
],
"due": "",
"letter": true,
"no": 16,
"note": "",
"owner": "기술이사",
"priority": "보통",
"proposer": "장수정",
"quick": "잘못된 사인물 사진·위치 정리",
"resolve": "시공 오류면 현산 하자로, 추가 요청이면 조합 공문 → 교체 완료 확인",
"status": "미착수",
"title": "305·307동 동입구 표시·엘리베이터 사인물 수정 교체",
"id": "a16"
},
{
"agenda": true,
"cat": "보안시설(CCTV·도어)",
"contacts": [
"스크린도어·볼라드 업체",
"조합",
"1·2단지 입대의",
"동대문소방서"
],
"due": "",
"letter": true,
"no": 17,
"note": "입구 두 곳 모두 소방도로. 1·2단지 비용분담 협의 중",
"owner": "기술이사 / 재무이사",
"priority": "높음",
"proposer": "정수, Jumbo",
"quick": "업체 2~3곳 현장실사·견적 요청, 설치 후보 위치 표시",
"resolve": "소방차 진입(비상 개방) 방식 소방서 협의 → 위치 확정 → 조합 통보 공문 → 1·2단지와 비용분담 합의 → 입대의 의결",
"status": "미착수",
"title": "출입구 스크린도어 + 소방도로 유압식 볼라드 사전작업",
"id": "a17"
},
{
"cat": "보안시설(CCTV·도어)",
"contacts": [
"관리사무소",
"도어 설치업체",
"동대문구청"
],
"due": "",
"letter": false,
"no": 18,
"note": "",
"owner": "기술이사",
"priority": "보통",
"proposer": "Jumbo",
"quick": "산책로 연결 지점 현장 사진, 외부인 출입 시간대 파악",
"resolve": "설치 필요성·방식 논의 → 견적 → 입대의 의결",
"status": "미착수",
"title": "천장산 산책로 도어 설치 검토",
"id": "a18"
},
{
"cat": "보안시설(CCTV·도어)",
"contacts": [
"조합",
"HDC현산",
"관리사무소",
"CCTV 업체"
],
"due": "",
"letter": true,
"no": 19,
"note": "기존 10대 증설 계획안 참고",
"owner": "기술이사",
"priority": "높음",
"proposer": "정수, 장수정",
"quick": "전기·통신·CCTV 배선도와 설치 공정표 요청 공문 발송",
"resolve": "도면 수령 → 사각지대·배선 검토 반영 → 공정표 확정 → 설치·준공 확인",
"status": "미착수",
"title": "CCTV 추가 설치 — 도면·공정표 확보",
"id": "a19"
},
{
"cat": "조경·외부공간",
"contacts": [
"조합",
"조경업체"
],
"due": "",
"letter": true,
"no": 20,
"note": "위치 확인 필요",
"owner": "기술이사",
"priority": "보통",
"proposer": "정수",
"quick": "정확히 어느 입구인지 입대의 내 확인, 현장 사진",
"resolve": "조경 준공도면(식재평면도) 확보 → 도면 대비 식재 여부 확인 → 조합에 추가식재 공문",
"status": "미착수",
"title": "입구 옹벽 조경수 확인(대나무 여부) 및 추가식재",
"id": "a20"
},
{
"cat": "조경·외부공간",
"contacts": [
"조합",
"조경업체"
],
"due": "",
"letter": true,
"no": 21,
"note": "",
"owner": "기술이사 / 재무이사",
"priority": "보통",
"proposer": "Jumbo",
"quick": "설치 위치·규모 정리",
"resolve": "당초 설계 포함 여부 조합 확인 → 미포함 시 견적·비용 주체 결정",
"status": "미착수",
"title": "개비온 옹벽 설치",
"id": "a21"
},
{
"cat": "조경·외부공간",
"contacts": [
"조합",
"조경업체"
],
"due": "",
"letter": true,
"no": 22,
"note": "301동 입주자 요청 포함",
"owner": "커뮤니티이사",
"priority": "보통",
"proposer": "Jumbo, 장수정",
"quick": "상태 불량 수목, 조명 영향받는 301동 위치 사진",
"resolve": "조합에 교체·추가식재 공문 → 식재 완료 확인. 조명 시간·각도는 #29 한예종 협의와 함께",
"status": "미착수",
"title": "한예종 라인 사철나무 교체 및 테니스장 조명 차단 식재",
"id": "a22"
},
{
"cat": "조경·외부공간",
"contacts": [
"HDC현산 CS팀",
"조경업체"
],
"due": "",
"letter": true,
"no": 23,
"note": "",
"owner": "기술이사",
"priority": "보통",
"proposer": "Jumbo",
"quick": "단지 내 고사목 위치 전수 조사(사진·번호)",
"resolve": "조경 하자보수 요청 → 교체 완료 확인",
"status": "미착수",
"title": "고사목 교체",
"id": "a23"
},
{
"cat": "조경·외부공간",
"contacts": [
"조합"
],
"due": "",
"letter": true,
"no": 24,
"note": "조합장님께 부탁드릴 일",
"owner": "커뮤니티이사",
"priority": "보통",
"proposer": "장수정",
"quick": "부지 경계·현황 사진",
"resolve": "조합장께 식재 요청 공문 → 식재 완료 확인",
"status": "미착수",
"title": "삼익공영 부지 식재 (사철나무 외 1~1.5m 수목)",
"id": "a24"
},
{
"cat": "조경·외부공간",
"contacts": [
"관리사무소",
"조합"
],
"due": "",
"letter": false,
"no": 25,
"note": "",
"owner": "재무이사",
"priority": "낮음",
"proposer": "Jumbo",
"quick": "펜스 필요 구간 표시, 견적 1~2곳",
"resolve": "비용 주체(조합/장충금/관리비) 결정 → 입대의 의결 → 설치",
"status": "미착수",
"title": "단지 내 잔디밭 펜스 설치",
"id": "a25"
},
{
"cat": "조경·외부공간",
"contacts": [
"조합",
"관리사무소"
],
"due": "",
"letter": true,
"no": 26,
"note": "",
"owner": "커뮤니티이사",
"priority": "보통",
"proposer": "장수정",
"quick": "입주민 아이디어 공모 (단톡방·설문)",
"resolve": "용도 확정 → 조합에 조경가구 설치 요청 공문",
"status": "미착수",
"title": "천장산 수변시설·잔디광장 활용 (용도 확정·조경가구 설치)",
"id": "a26"
},
{
"cat": "조경·외부공간",
"contacts": [
"조합",
"HDC현산"
],
"due": "",
"letter": false,
"no": 27,
"note": "",
"owner": "기술이사",
"priority": "낮음",
"proposer": "장수정",
"quick": "대상 위치 사진·설명 받기 (장수정 님 추후 공유)",
"resolve": "변경 사유 정리 → 조합/현산 요청",
"status": "미착수",
"title": "바닥 조경석 변경",
"id": "a27"
},
{
"cat": "하자·누수",
"contacts": [
"HDC현산 CS팀",
"관리사무소"
],
"due": "",
"letter": true,
"no": 28,
"note": "",
"owner": "기술이사",
"priority": "높음",
"proposer": "장수정",
"quick": "누수·침수 이력 세대·공용부 목록과 사진 취합",
"resolve": "현산에 일괄 하자 접수 → 보수 일정 회신 → 완료 확인 (하자담보기간 내)",
"status": "미착수",
"title": "누수·침수로 인한 마감 불량 하자 점검",
"id": "a28"
},
{
"cat": "대외 민원·협의",
"contacts": [
"한국예술종합학교"
],
"due": "",
"letter": true,
"no": 29,
"note": "",
"owner": "회장",
"priority": "보통",
"proposer": "Jumbo",
"quick": "소음 발생 시간·위치 기록, 영향 세대 파악",
"resolve": "한예종 시설 담당과 면담 → 작업시간·방음 조치 합의 (테니스장 조명 포함)",
"status": "미착수",
"title": "한예종 송추공방 및 기타 소음 협의",
"id": "a29"
},
{
"cat": "대외 민원·협의",
"contacts": [
"국가유산청(구 문화재청)",
"동대문구청"
],
"due": "",
"letter": true,
"no": 30,
"note": "문화재청은 2024년 국가유산청으로 개편",
"owner": "총무이사",
"priority": "보통",
"proposer": "장수정",
"quick": "사방댐 위치·관리 주체 확인, 과거 민원 이력 정리",
"resolve": "연간 점검 요청 시기·양식 표준화 → 관리사무소 정기 업무로 이관",
"status": "미착수",
"title": "사방댐 정기관리 요청 민원 매뉴얼화",
"id": "a30"
},
{
"cat": "대외 민원·협의",
"contacts": [
"동대문구청"
],
"due": "",
"letter": true,
"no": 31,
"note": "",
"owner": "총무이사",
"priority": "보통",
"proposer": "장수정",
"quick": "사진 첨부 설명 받기 (장수정 님 추후 공유)",
"resolve": "민원 접수 → 구청 현장 확인·조치 회신",
"status": "미착수",
"title": "천장산 배수 관련 구청 민원 (삼익공영 부지 경계)",
"id": "a31"
}
];
