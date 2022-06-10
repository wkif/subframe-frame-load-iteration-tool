def getHelp():
    filePath= './assets/help.txt'
    f = open(filePath, encoding='utf-8')
    txt = ''
    for line in f:
        txt+=(line.strip())+'\n'
    # print(txt)
    return txt
