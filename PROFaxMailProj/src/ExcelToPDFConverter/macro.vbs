Dim Excel
Dim ExcelDoc

Set Excel = CreateObject("Excel.Application")

'Open the Document
Set ExcelDoc = Excel.Workbooks.open("C:\Users\user\OneDrive\Musica\Cartella Rapporti Excel\CASSA DI RISPARMIO DI ASTI S.P.A. - CRASTI 3 ott 2020, 03 39 10.xlsx.xlsx")
Excel.ActiveSheet.ExportAsFixedFormat 0, "C:\Users\user\OneDrive\Musica\Cartella Rapporti PDF\CASSA DI RISPARMIO DI ASTI S.P.A. - CRASTI 3 ott 2020, 03 39 10.xlsx.pdf" ,0, 1, 0,,,0
Excel.ActiveWorkbook.Close
Excel.Application.Quit