#레드북 가사 공백제거용
#사용법: 가사 복붙 후 0 입력

name = input()
f= open("D:/Me/레드북가사/"+name+".txt","w+") #파일 생성

haha = []
while True:
    n = input()
    n = n.replace(u"\u200b", u"")
    if(n == "0"):
        break
    else:
        if(n != "" and n != " " and not "[" in n):
            haha.append(n)

for i in haha:
    f.write(i+"\n")
f.close()
