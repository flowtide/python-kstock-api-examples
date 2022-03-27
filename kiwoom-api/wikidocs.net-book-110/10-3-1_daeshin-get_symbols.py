# 종목 코드 가져오기
# https://wikidocs.net/3686
import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

# CetStockListByMarket 메서드의 인자로 1을 전달하면 유가증권시장의 종목을 파이썬 튜플 형태로 반환받을 수 있음
codeList = instCpCodeMgr.GetStockListByMarket(1)
nlines = 0
filename = 'kospi.csv'
f = open(filename, 'w')
for code in codeList:
    name = instCpCodeMgr.CodeToName(code)
    type = instCpCodeMgr.GetStockSectionKind(code)
    line = "%s,%s,%d" % (code, name, type)
    print(line)
    f.write("%s\n" % (line))
    nlines += 1
f.close()
print(f'{nlines} symbols written to "{filename}"')
