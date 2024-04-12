#가사 정리용 (등장인물 이름 및 불필요한 줄바꿈 없앰)
#사용법: 가사 복붙 후 0 입력

name = input()
f= open("D:/Me/가사/"+name+".txt","w+") #파일 생성

exceptList = ["다이애나", "댄", "나탈리", "게이브",'헨리','파인 박사','의사들'] # 제거할 등장인물 이름 넣으면 됨
haha = [] # 추출된 가사들
while True:
    line = input()
    line = line.replace(u"\u200b", u"")
    if(line == "0"):
        break
    
    if(line == "" or line == " "):
        continue
    
    flag = True
    for exc in exceptList:
        if exc in line:
            flag = False
            break
    if flag:
        haha.append(line)
        
for i in haha:
    f.write(i+"\n")
f.close()
