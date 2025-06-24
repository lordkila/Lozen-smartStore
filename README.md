# 스마트스토어 엑셀 일괄처리 프로그램  
Smart Store Excel Batch Processor

---

### 소개
스마트스토어 주문 데이터와 운송장 데이터를 자동 매칭하여 엑셀로 저장하는 PyQt5 기반 GUI 프로그램입니다. 암호화된 엑셀도 지원하며, 드래그 앤 드롭으로 파일을 올릴 수 있습니다.
 - [Click to read English introduction.](#English Guide)

   
### 주요 기능
- **드래그 앤 드롭**으로 주문/운송장 엑셀 업로드
- **암호화 엑셀 지원** (`msoffcrypto` 사용)
- **수취인명, 전화번호, 주소 기준 자동 매칭**
- **결과 테이블 미리보기 및 엑셀 저장**
- **헤더 스타일(폰트, 배경색 등) 자동 적용**

### 설치 방법
#### 필수 조건
- Python 3.7 이상
- pip

#### 의존성 설치
  ```bash
  pip install PyQt5 qasync pandas openpyxl msoffcrypto  
 ```

### 사용 방법
1. 프로그램 실행
python delivery_ui.py
2. 주문/운송장 엑셀 파일을 각각 드래그 앤 드롭(또는 클릭)하여 선택
3. "일괄처리 시작" 클릭
4. 매칭 결과가 테이블에 표시되고, 엑셀로 저장됨 (`일괄처리_[A파일명].xlsx`)
5. 저장된 파일을 열지 여부 안내

### 파일 포맷
#### 주문 데이터(A)
- 상품주문번호, 배송방법, 택배사, 상품명, 수량, 수취인명, 수취인연락처1, 통합배송지

#### 운송장 데이터(B)
- 수하인명, 수하인전화, 수하인주소1, 운송장번호

### 기술 정보
- **GUI**: PyQt5
- **비동기**: qasync
- **엑셀 암호 해독**: msoffcrypto
- **엑셀 스타일**: openpyxl (헤더 굵게, 흰색 폰트, 파란 배경, 가운데 정렬)

### 라이선스
GPL-3.0  
자세한 내용은 `LICENSE` 파일 참고


### 예시 폴더 구조
  ```bash
/SmartStore-Excel-Processor  
├── delivery_ui.py  
├── requirements.txt   
├── LICENSE  
└── README.md  
 ```

## 버전
- v1.0.0: 최초 배포
- v1.1.0: 암호화 엑셀 지원 추가
- v1.2.0: UI 개선 및 성능 최적화

---

## English Guide

### Introduction
A GUI tool for batch matching Smart Store order and shipping Excel files. Built with PyQt5, supports drag & drop, password-protected Excel, and saves results with styled headers.

### Features
- **Drag & Drop** upload for order/shipping Excel files
- **Encrypted Excel support** (`msoffcrypto`)
- **Auto-matching** by name, phone, and address
- **Result table preview and Excel export**
- **Styled headers** (font, background color, alignment)

### Installation
#### Requirements
- Python 3.7+
- pip

#### Install dependencies
  ```bash
  pip install PyQt5 qasync pandas openpyxl msoffcrypto  
 ```

### Usage
1. Run the program:
2. Drag and drop (or click) to select order and shipping Excel files
3. Click "Start Batch Processing"
4. Matching results are shown in the table and saved as Excel (`일괄처리_[A_filename].xlsx`)
5. Prompt to open the saved file

### File Formats
#### Order Data (A)
- 상품주문번호 (Order No), 배송방법 (Delivery), 택배사 (Courier), 상품명 (Product), 수량 (Qty), 수취인명 (Recipient), 수취인연락처1 (Phone), 통합배송지 (Address)

#### Shipping Data (B)
- 수하인명 (Recipient Name), 수하인전화 (Phone), 수하인주소1 (Address), 운송장번호 (Tracking No)

### Technical Details
- **GUI**: PyQt5
- **Async**: qasync
- **Excel Decryption**: msoffcrypto
- **Excel Styling**: openpyxl (bold header, white font, blue background, centered)

### License
GPL-3.0  
See `LICENSE` for details

### Example Structure
 ```bash
/SmartStore-Excel-Processor  
├── delivery_ui.py  
├── requirements.txt   
├── LICENSE  
└── README.md  
 ```

### version
v1.0.0: Initial release
v1.1.0: Added encrypted Excel file support
v1.2.0: UI improvements and performance optimization
