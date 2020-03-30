from __future__ import print_function
from datetime import date
from mailmerge import MailMerge
import xlrd

#name of the contract template in the same folder
contract_template = "template.docx"

#name of the exel sheet and worksheet containing the data
workbook = xlrd.open_workbook("Excel.xlsx")
worksheet = workbook.sheet_by_name("information")

contract = MailMerge(contract_template)

print(contract.get_merge_fields())

#tokens are the names of the tokens in the worksheets
token_one = worksheet.cell_value(1,1)
token_two= worksheet.cell_value(2,1)

#run everything
contract.merge(Test = man_name_short)
contract.write("merged_contract.docx")

