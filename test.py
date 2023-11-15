from openpyxl import load_workbook
from efc.interfaces.iopenpyxl import OpenpyxlInterface

wb = load_workbook('test.xlsx')
interface = OpenpyxlInterface(wb=wb, use_cache=True)
result = interface.calc_cell('B33', 'l1')
print(result)
exit()
# e.g. A1 stores formula '=1 + 2', then result = 3
for i in range(13, 20):
    result = interface.calc_cell(f'{chr(ord("A") + i)}28', 'l1')
    print(i, result, type(result))

exit(0)
result = interface.calc_cell('E28', 'l1')
print(result)
exit(0)
# EFC does not change the source document
print(wb['Worksheet1']['A1'].value)  # prints '=1 + 2'

# If you need to replace a formula in a workbook with a value,
# you need to do this
wb['Worksheet1']['A1'].value = interface.calc_cell('A1', 'Worksheet1')
print(wb['Worksheet1']['A1'].value)  # prints '3'

# The EFC does not track changes to values in the workbook.
# If the use_cache=True option is used, the calculated formulas
# are not recalculated again when they are accessed.
# e.g. A2 = 2, A3 = 1, A4 = A2 + A3
print(interface.calc_cell('A4', 'Worksheet1'))  # prints '3'
wb['Worksheet1']['A2'].value = 1234
print(interface.calc_cell('A4', 'Worksheet1'))  # prints '3'

# If you have made changes to the workbook, then you need to reset
# the cache to get up-to-date results
interface.clear_cache()
print(interface.calc_cell('A4', 'Worksheet1'))  # prints '1235'

# You can disable caching of results,
# but then when you run a large number of related formulas,
# the calculation speed will decrease significantly