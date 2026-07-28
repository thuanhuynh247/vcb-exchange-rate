from openpyxl import load_workbook

f = r'D:\Tygia-Tudong\TyGia_Banking_Temp.xlsx'
wb = load_workbook(f)
ws = wb['Dashboard']
for i, dv in enumerate(ws.data_validations.dataValidation):
    print(f'DV {i}: sqref="{dv.sqref}", type="{dv.type}", formula1="{dv.formula1}", showDropDown={dv.showDropDown}, showErrorMessage={dv.showErrorMessage}')
