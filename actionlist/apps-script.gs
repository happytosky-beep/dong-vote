/**
 * 이문아이파크자이 3단지 입대의 액션리스트 — 구글 시트 저장소 (Apps Script)
 * 1) 구글 시트를 새로 만들고 [확장 프로그램 → Apps Script]에 이 코드를 통째로 붙여넣기
 * 2) 아래 PIN_INIT 을 입대의 편집 비밀번호로 바꾸기
 * 3) 함수 선택에서 setup 실행 (권한 허용) → 시트에 31개 항목이 채워짐
 * 4) [배포 → 새 배포 → 웹 앱] 실행 사용자: 나 / 액세스: 모든 사용자 → 배포 → 웹 앱 URL 복사
 */
const PIN_INIT = '여기에_비밀번호';
const SHEET = 'items';
const COLS = ['id','no','cat','title','contacts','quick','resolve','owner','proposer','priority','status','letter','agenda','due','note','updatedAt'];

function sh_() { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function toRow_(o) {
  return COLS.map(k => {
    let v = o[k];
    if (k === 'contacts') v = Array.isArray(v) ? v.join(', ') : (v || '');
    if (k === 'letter' || k === 'agenda') v = v ? 'Y' : '';
    return v === undefined || v === null ? '' : String(v);
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
function readAll_() {
  const v = sh_().getDataRange().getValues();
  if (v.length < 2) return [];
  return v.slice(1).filter(r => r[0]).map(fromRow_);
}

function setup() {
  PropertiesService.getScriptProperties().setProperty('PIN', PIN_INIT);
  const s = sh_();
  s.getRange(1, 1, 2000, COLS.length).setNumberFormat('@');
  if (s.getLastRow() > 1) return;
  s.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  const now = new Date().toISOString();
  const rows = SEED.map(o => toRow_(Object.assign({ updatedAt: now }, o)));
  s.getRange(2, 1, rows.length, COLS.length).setValues(rows);
  s.setFrozenRows(1);
}

function doGet() { return json_({ ok: true, items: readAll_(), at: new Date().toISOString() }); }

function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: '잘못된 요청' }); }
  const pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin || String(b.pin || '') !== pin) return json_({ ok: false, error: '편집 비밀번호가 맞지 않습니다' });
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const s = sh_(); const v = s.getDataRange().getValues();
    const now = new Date().toISOString();
    if (b.action === 'add') {
      const maxNo = v.slice(1).reduce((m, r) => Math.max(m, Number(r[1]) || 0), 0);
      const it = Object.assign({}, b.item, { id: 'i' + Date.now(), no: maxNo + 1, updatedAt: now });
      s.appendRow(toRow_(it));
      return json_({ ok: true, item: it });
    }
    const idx = v.findIndex((r, i) => i > 0 && String(r[0]) === String(b.id));
    if (idx < 1) return json_({ ok: false, error: '항목을 찾지 못했습니다' });
    if (b.action === 'update') {
      const cur = fromRow_(v[idx]);
      const it = Object.assign(cur, b.item, { id: cur.id, no: cur.no, updatedAt: now });
      s.getRange(idx + 1, 1, 1, COLS.length).setValues([toRow_(it)]);
      return json_({ ok: true });
    }
    if (b.action === 'delete') { s.deleteRow(idx + 1); return json_({ ok: true }); }
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
"note": "[11/17 의견] 위탁업체 선정 시 북카페 화구시설(#33)·사우나(#34) 운영 여부를 함께 고려",
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
"note": "[11/17 의견] 커뮤니티 시설 인수인계 때 곰팡이 원인 제거를 확인한 뒤 인수 승인 (회장님)",
"owner": "회장 / 커뮤니티이사",
"priority": "높음",
"proposer": "장수정",
"quick": "발생 위치 사진, 환기설비 가동 상태 확인",
"resolve": "현산에 원인 조사·제거 요청 → 커뮤니티 인수인계(#8) 점검 항목에 포함 → 제거 확인 후 회장이 인수 승인",
"status": "미착수",
"title": "골프장·사우나 곰팡이 냄새 원인 제거",
"id": "a11"
},
{
"cat": "커뮤니티센터",
"contacts": [
"관리사무소",
"조합",
"HDC현산"
],
"due": "",
"letter": false,
"no": 12,
"note": "[11/17 의견] 게스트룸 인수인계 여부를 먼저 확인한 뒤 운영 비용 등 논의",
"owner": "커뮤니티이사",
"priority": "낮음",
"proposer": "장수정",
"quick": "게스트룸 현장 확인(비품·하자)",
"resolve": "인수인계 대상·시기 확인(조합·현산) → 운영비·이용료 산정 → 이용규정 마련 후 입대의 의결 → 오픈",
"status": "미착수",
"title": "게스트룸 상태 확인 및 운영 준비",
"id": "a12"
},
{
"agenda": true,
"cat": "주차·출입",
"contacts": [
"관리사무소",
"출입통제 시스템 업체",
"HDC현산 전기팀"
],
"due": "",
"letter": false,
"no": 13,
"note": "[11/17 의견] 시공사 협의 사항. 305동에서 307동 비밀번호가 안 되는 건 현산 전기팀이 어렵다고 했다고 알려져 있어 사실 확인 필요",
"owner": "기술이사",
"priority": "높음",
"proposer": "정수, 장수정",
"quick": "현산 전기팀에 305동 공용 비밀번호 출입이 기술적으로 불가한지 서면 확인 요청 (사유·대안 포함)",
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
"note": "[11/17 의견] 1가구 2대 이상 추가 주차요금은 주차관리규정 제정 때 함께 논의",
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
"관리사무소",
"동대문구청"
],
"due": "",
"letter": false,
"no": 15,
"note": "[11/17 의견] 주차관리규정 제정 때 논의. 법정 전기차 전용 주차면 수를 확인한 뒤, 남는 면에 일반차 병행주차가 가능한지 구청 확인 필요",
"owner": "기술이사",
"priority": "보통",
"proposer": "장수정",
"quick": "입주민 설문(전기차 보유·구매 예정) 배포",
"resolve": "법정 전기차 전용 주차면 수 확인 → 잔여 면적의 전기차·일반차 병행주차 가능 여부 구청 질의 → 주차관리규정(#14)에 반영해 의결 → 도색·표지 정비",
"status": "미착수",
"title": "전기차 수요조사 및 주차구획 정비",
"id": "a15"
},
{
"cat": "하자·누수",
"contacts": [
"HDC현산 CS팀",
"엘리베이터 유지보수업체",
"관리사무소"
],
"due": "",
"letter": true,
"no": 16,
"note": "[11/17 의견] 동입구 표시·엘리베이터 사인물은 시공사 협의 사항\n설계·시공 하자로 즉시 시정 필수. 기존 '305·307동 사인물 수정 교체' 항목을 확대",
"owner": "기술이사",
"priority": "높음",
"proposer": "장수정, 입대의 위원",
"quick": "동·호기별로 층 표시 오류, 안내음성 오류, 동선 혼선 지점을 사진·영상으로 목록화 (동입구 표시·사인물 포함)",
"resolve": "현산에 하자 시정 공문(시정 기한 명시) → 승강기 업체와 현장 확인 → 표시·음성 설정 수정, 동선 설계 오류는 개선안 회신 → 완료 확인. 미이행 시 #32와 묶어 조합 경유 압박",
"status": "미착수",
"title": "305·306·307동 엘리베이터 층별 안내표시·안내음성·동선 하자 즉시 시정 (동입구 표시 포함)",
"agenda": true,
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
"note": "[11/17 의견] 현재 CCTV 설치가 예정된 것으로 알려져 있어 시공사와 내용 확인 필요\n기존 10대 증설 계획안 참고",
"owner": "기술이사",
"priority": "높음",
"proposer": "정수, 장수정",
"quick": "현산에 예정된 CCTV 설치 범위(위치·대수)·일정 확인 요청, 도면·공정표 함께 요청",
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
"note": "[11/17 의견] 301동 입주자 요청: 한예종 테니스장 조명 차단용 추가 식재\n301동 입주자 요청 포함",
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
"note": "[11/17 의견] 입대의 논의로 활용안을 정리한 뒤 입주민 의견 수렴",
"owner": "커뮤니티이사",
"priority": "보통",
"proposer": "장수정",
"quick": "입대의에서 활용안 2~3개 정리",
"resolve": "입대의 정리안 → 입주민 의견함·설문으로 의견 수렴 → 용도 확정 → 조합에 조경가구 설치 요청 공문",
"status": "미착수",
"title": "천장산 수변시설·잔디광장 활용 (용도 확정·조경가구 설치)",
"id": "a26"
},
{
"cat": "조경·외부공간",
"contacts": [
"HDC현산",
"조합"
],
"due": "",
"letter": false,
"no": 27,
"note": "[11/17 의견] 시공사 협의 필요",
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
"note": "[11/17 의견] 하자보수 항목으로 처리",
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
"note": "[11/17 의견] 국가유산청(구 문화재청)에 공문으로 관리 주체·점검 주기 확인 필요\n문화재청은 2024년 국가유산청으로 개편",
"owner": "총무이사",
"priority": "보통",
"proposer": "장수정",
"quick": "국가유산청에 사방댐 관리 주체·정기점검 주기 질의 공문 발송",
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
"note": "[11/17 의견] 조합 승인 완료. 구청 민원까지 필요한지 확인 필요",
"owner": "총무이사",
"priority": "보통",
"proposer": "장수정",
"quick": "조합 승인 내용(조치 범위·일정) 확인 → 그것으로 충분한지 판단",
"resolve": "조합 조치로 해결되면 완료 확인, 부족하면 사진 첨부해 구청 민원 접수",
"status": "진행중",
"title": "천장산 배수 관련 구청 민원 (삼익공영 부지 경계)",
"id": "a31"
},
{
"no": 32,
"cat": "하자·누수",
"title": "304·306·307동 진입로 반복 침수 근본대책 (현산 6월 약속 미이행)",
"contacts": [
"HDC현산 CS팀",
"조합",
"관리사무소",
"동대문구청"
],
"owner": "기술이사 / 회장",
"proposer": "입대의 위원",
"priority": "높음",
"status": "미착수",
"letter": true,
"agenda": true,
"due": "",
"quick": "침수 일시·강수량·사진·영상 기록 정리, 현재 모래주머니 임시조치 현황, 6월 현산 약속 근거(회의록·문자) 확보",
"resolve": "현산에 방지대책 이행계획·공정표 회신 요구 공문(회신기한 14일) → 미회신 시 조합 경유 재요구·구청 민원, 필요하면 하자심사분쟁조정 검토 → 2027 우기 전 공사 완료·준공 확인",
"note": "단지 중앙통로 인접 304·306·307동 진입로가 많은 비에 반복 침수. 엘리베이터·전기 계통 고장, 주차장 천장 붕괴 위험. 현산이 대책 마련을 약속했으나 6월 이후 미이행. 최초 TFT 발족의 핵심 사유이므로 최우선 집중 대응 (#28 마감 불량 하자와 연계)",
"id": "a32"
},
{
"cat": "커뮤니티센터",
"owner": "커뮤니티이사",
"proposer": "입대의 위원",
"priority": "보통",
"status": "미착수",
"letter": false,
"agenda": false,
"due": "",
"no": 33,
"title": "북카페 화구시설(기존 조식서비스용) 운영·관리 방안",
"contacts": [
"커뮤니티 위탁운영업체",
"관리사무소"
],
"quick": "화구·배기 설비 현황과 인수 상태 확인, 인근 단지의 소규모 북카페·조식 운영 사례 수집",
"resolve": "위탁업체 선정(#9) 때 운영 가능 여부 질의 → 조식이 수지 안 맞으면 대안 비교(무인 카페·간편식 자판기, 시간제 외부 임대, 주민 공유주방·모임공간, 화구 폐쇄 후 북카페 전용) → 입주민 의견 수렴 → 입대의 의결",
"note": "[11/17 의견] 규모가 작아 조식서비스는 수지타산이 맞지 않아 업체가 운영을 거절할 가능성이 큼. 대안 의견 수렴 필요 (급하지 않음, 인수인계 후 위탁업체 선정 시 고려)",
"id": "a33"
},
{
"cat": "커뮤니티센터",
"owner": "커뮤니티이사",
"proposer": "입대의 위원",
"priority": "보통",
"status": "미착수",
"letter": false,
"agenda": false,
"due": "",
"no": 34,
"title": "사우나 운영·관리 방안 (소규모 단지 운영비 부담)",
"contacts": [
"커뮤니티 위탁운영업체",
"관리사무소"
],
"quick": "사우나 설비(보일러·순환·환기) 사양과 예상 운영비(가스·전기·수도·청소·위생) 자료 요청, 비슷한 규모 단지 사례 수집",
"resolve": "운영비·예상 이용자 수로 세대당 부담 산출 → 대안 비교(요일·시간 단축 운영, 유료화·수익자 부담, 샤워실만 운영, 휴관) → 입주민 의견 수렴 → 입대의 의결 후 위탁 조건에 반영",
"note": "[11/17 의견] 사우나는 운영비 부담이 매우 큰 시설이고 단지 규모가 작아 운영 방식 검토 필요. #11 곰팡이 문제와 연계 (급하지 않음, 인수인계 후 위탁업체 선정 시 고려)",
"id": "a34"
}
];
