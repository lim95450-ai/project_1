const fs = require('fs');
const path = require('path');

const SRC_DIR = __dirname;
const DIST_DIR = path.join(__dirname, 'dist');

// 빌드 대상 HTML 파일 목록
const TARGET_FILES = ['index.html', 'admin.html', 'view.html'];

function build() {
  console.log('🚀 Vercel 배포용 빌드 스크립트 실행 시작...');

  // 1. dist 디렉토리 관리
  if (fs.existsSync(DIST_DIR)) {
    fs.rmSync(DIST_DIR, { recursive: true, force: true });
    console.log('🧹 기존 dist 디렉토리를 삭제했습니다.');
  }
  fs.mkdirSync(DIST_DIR);
  console.log('📁 신규 dist 디렉토리를 생성했습니다.');

  // 2. 환경변수 확인
  const gasWebAppUrl = process.env.GAS_WEBAPP_URL || '';
  if (!gasWebAppUrl) {
    console.warn('⚠️ [경고] 환경변수 GAS_WEBAPP_URL이 설정되지 않았습니다.');
    console.warn('   Vercel 대시보드에서 환경 변수를 추가하거나 로컬에서 주입하여 빌드하세요.');
  } else {
    console.log(`🔗 설정된 Apps Script URL: ${gasWebAppUrl}`);
  }

  // 3. 파일 처리
  for (const filename of TARGET_FILES) {
    const filePath = path.join(SRC_DIR, filename);
    if (!fs.existsSync(filePath)) {
      console.warn(`⚠️ [경고] 파일이 존재하지 않아 건너뜁니다: ${filename}`);
      continue;
    }

    console.log(`📦 처리 중: ${filename}`);
    let content = fs.readFileSync(filePath, 'utf8');

    // 3-1. <?!= include('파일명') ?> 구문을 해당 파일의 내용으로 치환
    const includeRegex = /<\?!=\s*include\(['"]([^'"]+)['"]\)\s*\?>/g;
    content = content.replace(includeRegex, (match, includeName) => {
      const includePath = path.join(SRC_DIR, `${includeName}.html`);
      if (fs.existsSync(includePath)) {
        console.log(`   └─ 인클루드 삽입: ${includeName}.html`);
        // 가져온 내용에 대해서도 재귀적으로 include가 있을 수 있으므로 단순 텍스트 삽입
        return fs.readFileSync(includePath, 'utf8');
      } else {
        console.error(`   ❌ [오류] 인클루드 파일이 존재하지 않습니다: ${includeName}.html`);
        return `<!-- ERROR: Include not found: ${includeName} -->`;
      }
    });

    // 3-2. 환경변수를 클라이언트에 바인딩하기 위해 헤더 영역에 window.GAS_WEBAPP_URL 전역 변수 정의 코드 삽입
    const envScript = `\n  <script>window.GAS_WEBAPP_URL = "${gasWebAppUrl}";</script>\n`;
    if (content.includes('</head>')) {
      content = content.replace('</head>', `${envScript}</head>`);
    } else {
      content = envScript + content;
    }

    // 3-3. Apps Script 전용 스크립틀릿 예외 제거 (예: view.html의 scheduleData 주입 코드 제거)
    // Vercel 환경에서는 클라이언트 fetch로 조회하므로, 기존 서버사이드 주입문 var scheduleData = <?!= scheduleData ?>; 은 에러를 내지 않게 주석 처리 또는 초기화
    content = content.replace(/var\s+scheduleData\s*=\s*<\?!=\s*scheduleData\s*\?>\s*;/g, 'var scheduleData = null; // GAS Server-side render disabled');

    // 3-4. 결과물 저장
    const destPath = path.join(DIST_DIR, filename);
    fs.writeFileSync(destPath, content, 'utf8');
    console.log(`   ✅ 빌드 완료 -> dist/${filename}`);
  }

  console.log('🎉 빌드 프로세스가 모두 성공적으로 끝났습니다!');
}

build();
