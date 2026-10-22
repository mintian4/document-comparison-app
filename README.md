# 문서 변경 비교기 웹사이트 배포 준비 파일

이 폴더의 `app.py`와 `requirements.txt`를 GitHub 저장소에 올린 뒤 Streamlit Community Cloud에서 배포하면 다른 사람이 인터넷 주소로 앱을 열 수 있습니다.

## 공개 배포 순서

1. GitHub 계정이 없다면 [github.com](https://github.com/)에서 계정을 만듭니다.
2. GitHub에서 새 저장소(repository)를 만듭니다. 공개 사이트로 배포하려면 저장소를 **Public**으로 설정합니다.
3. 이 폴더의 `app.py`, `requirements.txt`, `README.md` 세 파일을 저장소에 올립니다. `.venv` 폴더, 문서 샘플이 아닌 실제 문서, API 키는 올리지 마세요.
4. [Streamlit Community Cloud](https://share.streamlit.io/)에 GitHub 계정으로 로그인하고 **Create app**을 누릅니다.
5. 방금 만든 저장소, `main` 브랜치, `app.py` 파일을 선택한 뒤 **Deploy**를 누릅니다.
6. 배포가 끝나면 `...streamlit.app` 주소가 생깁니다. 그 주소를 다른 사람에게 보내면 사용할 수 있습니다.

공개 GitHub 저장소의 코드는 누구나 볼 수 있습니다. Streamlit 문서에 따르면 Community Cloud 앱은 저장소 공개 여부를 따르며, 공개 저장소의 앱은 기본적으로 공개 앱입니다. 따라서 이 폴더에는 앱 코드와 패키지 목록만 두었습니다.

## 문서 개인정보 안내

인터넷 공개 배포를 하면 사용자가 올린 문서는 내 노트북이 아니라 호스팅 서버에서 처리됩니다. 이 앱은 문서 비교 자체에는 외부 AI를 사용하지 않지만, AI 요약을 켜고 API 키를 설정하면 비교 텍스트를 OpenAI API로 전송합니다. 민감 정보, 개인정보, 실제 GMP 문서는 업로드하지 말고 가상 문서로만 시연하세요.

## 제한 사항

- 비교 기능은 문서에서 추출한 글자를 줄 단위로 비교합니다.
- 이미지로 스캔한 PDF는 OCR이 없어 글자를 읽지 못할 수 있습니다.
- HWP의 표, 글상자, 수식은 정확히 읽히지 않을 수 있습니다. 가능하면 HWPX로 저장한 파일을 사용하세요.
- 이 교육용 앱은 공식 문서관리·품질 승인 도구가 아닙니다.
