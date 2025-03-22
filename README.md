❗엑셀로 작성한 테이블 데이터를 C#에서 사용하기 위한 cs/json으로 변환해주는 프로그램입니다.

변환하고 싶은 xlsx 파일 리스트를 txt파일로 작성후 경로를 지정해주시면 변환이 가능합니다.
![image](https://github.com/user-attachments/assets/5b8a6834-54d0-4b14-9ea7-94427b6582ac)

![image](https://github.com/user-attachments/assets/f6017fdb-a4e4-4f7a-b9c0-ce398d5b7024)

❗Enum 타입의 변환.

열거형 데이터를 변환하는 경우 엑셀에서 타입 이름 앞에 접두사 #E_를 붙여야 변환이 가능합니다.
![enum](https://github.com/user-attachments/assets/658073a0-6807-437d-bbc9-1b910ba19cbc)

❗cs 파일 커스텀

변환되는 cs파일의 내용을 변경하고 싶으시다면 bin/Debug/net8.0-windows 파일 내의 ClassTemplate.txt 파일의 내용을 수정하시면 됩니다.
