# students_homework_checker
학생 숙제관리 프로그램 

### 배경 
- 네이버 밴드 내 각 숙제 게시물에 댓글로 숙제 제출 
- 학생이 각 숙제 별로 제출했는지 확인하는 시트 생성 의뢰

### 상세 기능
- 댓글과 댓글작성자 수집
- 퀴즈의 경우 엑셀 다운로드로 수집
- 해당 주차 제출 현황 JOIN
- 이전 주차 제출 현황과 JOIN 
- 숙제 미제출 시 초기 부여된 포인트 차감 코드 추가 

### file
- [homework.py](https://github.com/hmii/students_homework_checker/blob/83d3fbcc1a1b4c3013af89658bdd62674acb5c10/homework.py)
  - 학생 목록 생성, 스크래핑, 병합, 제출 현황 등
- [makedic.py](https://github.com/hmii/students_homework_checker/blob/83d3fbcc1a1b4c3013af89658bdd62674acb5c10/makedic.py)
  - 반별 밴드 주소 딕셔너리 생성
- [use_module.ipynb](https://github.com/hmii/students_homework_checker/blob/83d3fbcc1a1b4c3013af89658bdd62674acb5c10/use_module.ipynb)
  - 실행 노트북 
