// 수업자료일람 저장소 연결 설정 (구글 시트 + Apps Script)
//  - appsScriptUrl : Apps Script를 [배포 > 새 배포 > 웹 앱]으로 배포한 뒤 나오는 "웹 앱 URL"
//                    (https://script.google.com/macros/s/.../exec 모양)
//  - sheetUrl      : (선택) 그 구글 시트의 주소. 넣으면 관리자 로그인 때 위 메뉴에 "구글 시트 열기"가 생깁니다.
// 두 값 모두 공개되어도 되는 주소입니다. 저장은 Apps Script가 비밀번호로 막습니다.
window.SITE_CONFIG = {
  appsScriptUrl: "",
  sheetUrl: ""
};
