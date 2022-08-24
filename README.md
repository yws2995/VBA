# VBA


#### 오류 무시하기
On Error Resume Error

#### 속성
- Value 셀에 입력할 수식을 문자열 형태로 지정
- Formula 수식 입력 가능
- FomulaR1C1 상대 참조 수식을 작성할 때 주로 사용
- Text 셀에 적용된 셀 서식을 적용한 문자열로 반환

#### 변수(변수명 + As 데이터형)   
숫자 변수는 초기값으로 0을 갖고 문자 변수는 공백("")으로 개체변수는 nothing(Empty)로 초기값을 갖는다. 
- Dim
- Static
- Private
- Public

주로 Dim, Public으로 선언하며 저장할 값의 종류와 크기에 따라 다르게 지정한다. VBA의 경우 데이터형을 지정하지 않으면 Variant형(16Byte)으로 지정한다.  

Variant형은 모든 데이터 형이 가능하지만 크기가 크다는 단점이 있는 데이터 형이다.   

Object라는 일반적인 개체 형태로도 데이터를 지정할 수 있으나 개체변수(Range, worksheet등)에 값을 지정할 때는 반드시 Set 문을 사용해야 한다.   
일반 데이터 타입의 사용목적 => 데이터의 입력/출력
개체 데이터타입의 사용 목적 => 데이터의 입력 / 출력을 위한 프로시저(명령문)으로 동작
Dim(Declare in memory) : 메모리 할당   
Set : 개체변수를 할당 
```
Dim MyName as String
Dim MyPic as Image

MyName = "이름"
Set MyPic = Sheet(1).Image(1)
```
#### 상수 ( Const + As 데이터형 )   
- 변수와 상수의 차이 : 변수는 값이 변경 가능한 메모리 할당공간이고, 상수는 값이 변경되지 않는 메모리 할당공간이다. 

#### 대화상자 ( MsgBox, Inputbox, Application.Inpiutbox )   
- MsgBox 메세지내용[.단추 종류 + 아이콘 종류 + 기본 단추의 위치, 제목]     

- MsgBox( 메세지내용[.단추 종류 + 아이콘 종류 + 기본 단추의 위치, 제목]  )   
=> - MsgBox의 반환 단추를 확인하기 위해 괄호로 감싸서 표현   

#### 오류 처리하기
1. On Error GoTo 레이블명
2. On Error Resume Next
3. On Error GoTo 0

#### For문

For 카운터변수 = 시작수 To 끝수[Step 증감값]
'실행할 내용들1
[Exit For]
'실행할 내용들 2
Next [카운터변수]
- step 증감값을 생략하면 step1로 지정
- 카운터 변수는 생략 가능

#### For Each문 (개체를 처리할 때는 For 문보다 효과적이다)
Dim 개체변수 as 개체형
For Each 개체변수 in 컬렉션개체
  '실행할 내용들1
  [Exit For]
  '실행할 내용들2
 Next [개체변수]
 
 #### 사용자 정의 폼 UserForm
 폼을 화면에 표시하는 방법은 모달과 모덜리스
 - 폼 표시
 - - 폼개체명.Show
 - 폼 닫기
 - - UNload 폼개체명
 
