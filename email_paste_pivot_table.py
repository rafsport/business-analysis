import win32com.client as win32

xl = win32.Dispatch('Excel.Application')
xlsx = xl.Workbooks.Open(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Dashboard inflow - Raffaele.xlsb")
xlsx.Sheets.Item('TuttotelFE').PivotTables('Tabella pivot1').TableRange2.Copy()

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Raffaele.Sportiello@wolterskluwer.com'
mail.Subject = 'Prova'
mail.Body = 'Message body'

mail.Display()

inspector = outlook.ActiveInspector()
word_editor = inspector.WordEditor
word_range = word_editor.Application.ActiveDocument.Content
word_range.PasteExcelTable(False, False, True)
