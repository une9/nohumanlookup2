# no human lookup 👀❌
귀찮아서 더 귀찮은 일 만들기
<br/><br/>
\-
<br/><br/>

**기존 엑셀 문서에 작성된 테이블/컬럼 정보와 실제 DB의 컬럼 정보가 일치하는지 비교해주는 프로그램** <br/><br/>

#### 🤷‍♀️ 5W1H
>**WHEN** : 2023년 7-8월 어느 더운 여름날들 <br/>
**WHAT** : 인터페이스 명세서 작업 중... <br/>
**WHY** : 차세대 DB 스키마 변경사항 확인을 해줘야 해서 <br/>
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

