# 과거 데이터 구하기
# https://wikidocs.net/3684
import win32com.client

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

instStockChart.SetInputValue(0, "A003540")
instStockChart.SetInputValue(1, ord('2'))
instStockChart.SetInputValue(4, 10)
instStockChart.SetInputValue(5, 5)
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()
numData = instStockChart.GetHeaderValue(3)
