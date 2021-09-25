from bs4 import BeautifulSoup
import requests
import pandas as pd
import os

# 년도
year = '2021'
# 학기(1학기 : 1, 2학기 : 2, 여름계절학기 : 5, 겨울계절학기 : 6)
hakgi = '2'

# 학수번호(이름 끝에 붙은 1,2,3,4는 학년을 의미)
# 예시
haksu1 = ['130819', '124201']
haksu2 = ['130411', '130412']
haksu3 = ['130611', '130613']
haksu4 = ['130820', '130813']

print("""
사전 입력된 학수번호를 기반한 {}-{}학기 강의의 모든 정보를 얻습니다.
""".format(year,hakgi))
print("-----------------------------------------------------------------------------")

# 바꿀 수 있는 값 : 년도, 학기, 학수번호
# 그 외의 코드는 수정 x
# 엑셀 파일이 열린 상태로 실행할 경우 "PermissionError" 오류가 발생하니 주의

def urlreturn(year,hakgi,haksu,bunban):
    return "https://sugang.hongik.ac.kr/cn13061.jsp?yy={}&hakgi={}&haksu={}&bunban={}".format(year,hakgi,haksu,bunban)

def get_res(haksu11, bunban11, grade1):
    haksu = haksu11
    bunban = bunban11

    url = urlreturn(year, hakgi, haksu, bunban)
    request = requests.get(url)
    html = request.text.strip()

    dict = {}
    error = 0
    soup = BeautifulSoup(html, 'html.parser')
    contents = list(soup.stripped_strings)

    if len(contents) == 0:
        error -= 2

    for i in range(len(contents)):
        if contents[i] == "개설학기":
            dict["개설년도/학기"]=contents[i+2]

        if grade1>0:
            dict["학년"]=str(grade1)

        if contents[i] == "교과목명":
            if contents[i+2] == "학수번호":
                dict["교과목명"] = "입력되지 않음"
            else:
                dict["교과목명"]=contents[i+2]

        if contents[i] == "학수번호":
            dict["학수번호"]=contents[i+2][:6]
            dict["분반"]=contents[i+2][len(contents[i+2])-1]

        if contents[i] == '학점/시수(이론/실기)':
            if contents[i+2] == "설계학점":
                dict["학점/시수(이론/실기)"] = "입력되지 않음"
            else:
                dict["학점/시수(이론/실기)"]=contents[i+3]

        if contents[i] == "강의시간":
            if contents[i+2] == "개설학과":
                dict["강의시간"] = "입력되지 않음"
            else:
                dict["강의시간"]=contents[i+2]

        if contents[i] == "강의실":
            if contents[i+2] == "담당교수":
                dict["강의실"] = "입력되지 않음"
            else:
                dict["강의실"] = contents[i+2]

        if contents[i] == "담당교수":
            dict["담당교수"]=contents[i+2]

        if contents[i] == "E-MAIL":
            if contents[i+1] == "연구실 및 면담시간":
                dict["이메일주소"] = "입력되지 않음"
            else:
                dict["이메일주소"]=contents[i+1]

        if "담당교수" not in contents:
            error += 1

    if error > 0:
        # 해당되는 강좌가 없거나, 아직 입력되지 않았습니다.
        return [0,0,0]
    if error < 0:
        # 적절하지 않은 값을 입력했거나 서버 문제 또는 오류가 발생하였습니다.
        return [0,0,0]
    if error == 0:
        # 정상적인 경우입니다.
        # 아래 주석처리된 코드는 테스트용
        # for j in dict:
            # print(j, ":", dict[j])
        # print(url)
        return [dict,url,1]

def grade():
    haksu1_res = []
    haksu2_res = []
    haksu3_res = []
    haksu4_res = []

    # 1학년 코드
    for l in haksu1:
        n = 1
        while True:
            if get_res(l,str(n),1)[2] == 1:
                haksu1_res += [[[get_res(l,str(n),1)[0]],[get_res(l,str(n),1)[1]]]]
                n += 1
            else:
                break

    # 2학년 코드
    for i in haksu2:
        n = 1
        while True:
            if get_res(i,str(n),2)[2] == 1:
                haksu2_res += [[[get_res(i,str(n),2)[0]],[get_res(i, str(n),2)[1]]]]
                n += 1
            else:
                break

    # 3학년 코드
    for j in haksu3:
        n = 1
        while True:
            if get_res(j,str(n),3)[2] == 1:
                haksu3_res += [[[get_res(j, str(n),3)[0]], [get_res(j, str(n),3)[1]]]]
                n += 1
            else:
                break

    # 4학년 코드
    for k in haksu4:
        n = 1
        while True:
            if get_res(k,str(n),4)[2] == 1:
                haksu4_res += [[[get_res(k, str(n),4)[0]], [get_res(k, str(n),4)[1]]]]
                n += 1
            else:
                break

    return haksu1_res, haksu2_res, haksu3_res, haksu4_res

sec = round(0.8 * len(haksu1+haksu2+haksu3+haksu4),0)
print("서버로부터 값을 받아오고 있습니다. 약 {}초 가량 소요됩니다.".format(sec))
print('-----------------------------------------------------------------------------')

# 딕셔너리를 데이터프레임으로 바꿈
df = pd.DataFrame()
# 강의계획서 링크 부분을 위의 데이터프레임에 붙이기 위한 작업
df1 = pd.DataFrame()
for i in grade():
    for j in i:
        df = df.append(j[0])
        df1 = df1.append(j[1])

df1.columns=["강의계획서 링크"]
df = pd.concat([df,df1],axis=1)

df = df.reindex(columns=["개설년도/학기","학년","교과목명","학수번호","분반","학점/시수(이론/실기)","강의시간","강의실","담당교수","이메일주소","강의계획서 링크"])
df.to_excel('./timetable.xlsx')

print(df)
print('-----------------------------------------------------------------------------')
print("엑셀 파일로 데이터를 export 하였습니다.")
print('-----------------------------------------------------------------------------')
os.system('pause')