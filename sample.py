import MatrixSpreadsheet as ms
m = ms.MatrixSpreadsheet('your_email','your_passwd','NameofExistingSpreadSheet')
data = [[1,2],[3,4]]
m.Store(data,'Sheet 1',(1,3))
new_data = m.Read('Sheet 1',position_=(1,3),rows_=2)
assert (new_data == data).all()
