GSmatrix
========

Move 2D data structures between Python and Google spreadsheet


Example
======

```python
import MatrixSpreadsheet as ms
m = ms.MatrixSpreadsheet('your_email','your_passwd','NameofExistingSpreadSheet')
data = [[1,2],[3,4]]
m.Store(data,'Sheet 1',(1,3))
new_data = m.Read('Sheet 1',position_=(1,3),rows_=2)
assert (new_data == data).all()
```


In details:

Import the module

```import MatrixSpreadsheet as ms```

Initialize the object with your username and password and the name of an existing worksheet.

```m = ms.MatrixSpreadsheet('your_email','your_passwd','NameofExistingSpreadSheet')```

Create sample data. At the moment both numeric and string data can be written, but only numeric data can be read.

```data = [[1,2],[3,4]]```

Store the data in the (new or existing) sheet named 'Sheet 1', starting at row number 1 and column 3.

```m.Store(data,'Sheet 1',(1,3))```

Read 2 rows of data from sheet named 'Sheet 1', starting
at row number 1 and column 3. The number of columns will be automatically determined to be the maximum sensible.

```new_data = m.Read('Sheet 1',position_=(1,3),rows_=2)```

And be surprised:

```assert (new_data == data).all()```


----------

See the test file for more usage examples.

Contributions (pull requests) welcome.
See Issues.
