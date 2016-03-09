import openpyxl
import warnings
from GetValueOfArea import GetValueOfArea

warnings.simplefilter("ignore")

wb = openpyxl.load_workbook(filename='СТГ-568-НС-ОМ_Предписания_01.xlsx')
sheet = wb['СКЗ']

val = GetValueOfArea(sheet,4,4)

print(val)
