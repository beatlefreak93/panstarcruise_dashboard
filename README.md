# NEOHELIOS CRUISE - 객실 현황 대시보드

잔여 객실 및 승객 현황을 실시간으로 조회할 수 있는 대시보드입니다.

## 🚀 Streamlit Community Cloud 배포 방법

### 1단계: GitHub 저장소 생성

1. [GitHub](https://github.com)에 로그인
2. 새 저장소 생성 (New repository)
3. 저장소 이름: `neohelios-cruise-dashboard` (원하는 이름으로 변경 가능)
4. **Private** 저장소로 설정 (DB 정보 보호)

### 2단계: 코드 업로드

```bash
cd C:\Users\Lenovo\Desktop\NEOHELIOS_CRUISE
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/[YOUR_USERNAME]/[YOUR_REPO_NAME].git
git push -u origin main
```

### 3단계: Streamlit Community Cloud 배포

1. [share.streamlit.io](https://share.streamlit.io) 접속
2. GitHub 계정으로 로그인
3. "New app" 클릭
4. 저장소 선택: `[YOUR_USERNAME]/[YOUR_REPO_NAME]`
5. Main file path: `독립_대시보드_앱.py`
6. **Advanced settings** 클릭 → **Secrets** 탭
7. 아래 내용 붙여넣기:

```toml
[database]
server = "neohelios-prod.database.windows.net"
base_database = "neohelios_base"
cruise_database = "neohelios_cruise"
username = "panstar_viewer"
password = "Vostmxk0712!"
```

8. "Deploy!" 클릭

### 4단계: 완료! 🎉

- 몇 분 후 앱이 배포됩니다
- URL이 생성됩니다: `https://[YOUR_APP_NAME].streamlit.app`
- 이 URL을 팀원들과 공유하세요!

## ⚠️ 중요 사항

### DB 접속 권한 확인

Azure SQL Server가 Streamlit Cloud의 IP에서 접속을 허용하는지 확인하세요:

1. Azure Portal → SQL Server → Networking
2. "Public network access" → Enabled
3. Firewall rules → "Allow Azure services and resources to access this server" 체크

또는 Streamlit Cloud의 IP 대역을 허용해야 할 수 있습니다.

### Private 저장소 권장

- DB 연결 정보가 포함되어 있으므로 **Private 저장소** 사용을 강력히 권장합니다
- Streamlit Community Cloud는 Private 저장소도 무료로 배포 가능합니다

## 🛠️ 로컬 개발

```bash
pip install -r requirements.txt
streamlit run 독립_대시보드_앱.py
```

## 📁 파일 구조

```
NEOHELIOS_CRUISE/
├── 독립_대시보드_앱.py        # 메인 애플리케이션
├── requirements.txt          # Python 패키지
├── packages.txt              # 시스템 레벨 패키지
├── .streamlit/
│   ├── config.toml          # Streamlit 설정
│   └── secrets.toml         # DB 정보 (로컬용, GitHub에 올리지 않음)
├── .gitignore               # Git 제외 파일
└── README.md                # 이 파일
```

## 🔒 보안

- `.streamlit/secrets.toml`은 GitHub에 업로드되지 않습니다 (`.gitignore`에 포함)
- Streamlit Cloud의 Secrets 기능을 사용하여 안전하게 DB 정보를 관리합니다
- 절대 DB 비밀번호를 코드에 직접 작성하지 마세요

## 📞 문의

문제가 있으면 개발자에게 연락하세요.

