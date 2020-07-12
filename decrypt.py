import msoffcrypto
import sys
import xlrd

if len(sys.argv)<4:
 print("Usage python decrypt.py <encrypted_file> <decrypted_file> <password>")
else:
 file = msoffcrypto.OfficeFile(open(sys.argv[1], "rb"))
 file.load_key(password=sys.argv[3])
 file.decrypt(open(sys.argv[2], "wb"))
 book=xlrd.open_workbook(sys.argv[2])
 sheet=book.sheet_by_index(0)
 cell = sheet.cell(0,0)
 print cell.value
