# no human lookup 👀❌
귀찮아서 더 귀찮은 일 만들기
<br/><br/>
\-
<br/><br/>

**기존 인터페이스 명세서 문서에 작성된 테이블/컬럼 정보와 실제 DB의 컬럼 정보가 일치하는지 비교해주는 프로그램** <br/>

#### 📈 ver2 레벨업 배경
오픈 준비 해야 하는데 차세대 쪽에서 변경사항이 대응개발 쪽으로 공유가 잘 안되는 것 같아서, DB 변경 스크립트 취합하면서 확인하는 김에 다시 한번 운영DB 기준으로 전부 돌려보면 좋을 것 같았다

#### ✨ 주요 변경사항
- ver1은 각 파일마다 비교 탭을 생성했다면, ver2는 모든 명세서를 비교한 결과를 한눈에 확인할 수 있도록 하나의 문서로 취합함 
  - 하드코딩되어있던 row 값 동적으로 변경
  - 작업하고 난 엑셀 파일은 닫아서 프로그램 종료 후 비교문서 하나만 열려있도록 정리
  - 각 파일로 이동하여 상세 내용을 쉽게 확인할 수 있도록 하이퍼링크 추가

<br/><br/>

#### 🤷‍♀️ 5W1H
>**WHEN** : 2023년 12월 어느 추운 겨울날들 <br/>
**WHAT** : 통합 테스트 및 인수인계 진행 중... <br/>
**WHY** : 차세대 DB 스키마 변경사항 확인을 해줘야 할 것 같아서 <br/>
**WHERE** : 집과 사무실에서 <br/>
**WHO** : 내가 <br/>
**HOW** :  chatGPT와 구글의 조언을 참고하여<br/>


<br/><br/>

## 💻 환경 설정 & 실행


### ✨ 개발환경
- Python3 (3.9.9/3.11.4)
- vscode
<br/><br/>


### ⚙ 실행 환경 세팅

#### [DB Connection Properties]

  - create `db.properties` file on root directory and set db properties
    
      ```properties
      # db credentials
      user=root
      passwd=mypassword1234
      host=localhost
      db=myschema
      charset=utf8
      ```

  #### [Target Excel Files]
  - 검사를 원하는 엑셀 파일들을 `/excels` 폴더 안에 넣어주세요
  - 각 엑셀 파일은 정해진 양식을 따라 작성되어 있어야 검사가 가능합니다
    - 차세대 인터페이스 명세서 양식에 따라 작성된 파일만 가능합니다
    - 테이블명이 작성된 열의 `N`번째 컬럼에는 `TABLE` 이 적혀 있어야 합니다
    - 검사할 엑셀 시트는 `세번째`에 위치해야 합니다
  - 검사가 종료되면 파일에 생성된 `비교` 탭에서 색이 칠해진 셀을 확인해주세요
    - 기존 문서와 다른 내용의 셀은 `[빨간색]`
    - 기존 문서에는 없는 내용이 추가된 셀은 `[노란색]`

#### [Install Libraries]
- Open Project with VSCode
- `` Ctrl + ` `` (open terminal)
- `$ python -m venv venv` (create virtual environment named *venv*)
- `$ source venv/Scripts/activate` (activate virtual environment - for Windows)
- `$ pip install -r requirements.txt` (install required dependencies)
  

<br/>

### 🔨 실행 (vscode)
- `$ source venv/Scripts/activate` (이미 활성화되어 있는 경우 생략)
- `$ python main.py [sample.xlsx]` (run)
  - 특정 파일만 검사하고 싶다면 해당 파일의 이름을 파라미터로 넣어주세요
  - 파라미터를 생략하면 `excels` 폴더 내의 모든 파일을 검사합니다



<br/><br/><br/>
<br/><br/><br/>

