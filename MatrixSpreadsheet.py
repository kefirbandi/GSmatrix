# -*- coding: utf-8 -*-
import gdata.spreadsheet.service
import gdata.service
import gdata.spreadsheet
import numpy
import time
import sys
from collections import defaultdict

class MatrixSpreadsheet:
    """Facilitates moving data between Google Spreadsheets and numpy 2-d arrays."""
    
    def __init__(self, email_, password_, spreadsheetName_=None, spreadsheetKey_ = None ):
        """Constructor for the MatrixSpreadsheet object.

        Takes an email, password corresponding to a gmail account and the name of
        a spreadsheet.

        Args:
          email_: [string] The e-mail address of the account to use for the sample.
          
          password_: [string] The password corresponding to the account specified by
            the email parameter.

          spreadsheetName_: [string] The name of the spreadsheet to use.

          spreadsheetKey_: [string] The document key of the spreadsheet to use. (This is the long identifier you see in the url.)
          Either spreadsheetName_ or spreadsheetKey_ has to be given, but not both.
        """
        
        self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
        self.gd_client.email = email_
        self.gd_client.password = password_

        self.gd_client.ProgrammaticLogin()

        if spreadsheetName_ and spreadsheetKey_ :
            raise Exception("Give either key or name, but not both")
        self.key = spreadsheetKey_ or self._KeyFromSpreadsheetName(spreadsheetName_)

        self.ws_map = {}
        #ws_map is our internal storage of the worksheet data
        #It is a dictionary, with worksheet names as the keys, the values are fixed-length lists as follows:
        #[0] : ID as needed by the API
        #[1] : Timestamp of last synchronisation with server
        #[2] : The dictionary object, storing the data
        #[3] : The worksheet object

        f = self.gd_client.GetWorksheetsFeed(self.key)
        
        for fs in f.entry:
             self._addToMap(fs)
            

    def _addToMap(self,wsf_):
        "Put the worksheet object wsf_ in our database"
        self.ws_map[wsf_.title.text] = [wsf_.id.text.split('/')[-1], #id as needed by the API
                                            0,#time when our local cache was updated
                                            {},#our local cache
                                            wsf_] #the object itself, may come handy later 
    
    def _cacheWS(self, worksheetName_, create_ = False, cellFeedOnly_ = False ):
        """Ensures that we cache the latest snapshot of the worksheet.
        
        Compares the timestamp of the cache with the timestamp of the on-line version,
        reloads cache if needed. Cache is stored in self.ws_map[worksheetName_][2]
        and also returned.
        
        Args:
          worksheetName_: [string] Name of the worksheet to reload.
          create_: [bool] Whether to create the worksheet if not found or rather to bowl a googly
          cellFeedOnly_: [bool] If False: create the cell dictionary (if modified since cached), put it in the cache and return it
              If True: do not create the dictionary, only return the gdata cellfeed object
          
        Returns:
          Worksheet cache
          
        Raises:
          Something if worksheet does not exist and create_ == False
        """
        try:
            wsf = self.gd_client.GetWorksheetsFeed(self.key,self._WorksheetIDFromName(worksheetName_))
        except (KeyError):
            #so it was not in our database
            #we need to check if it exists?
            wsfs = self.gd_client.GetWorksheetsFeed(self.key)
            for wsf in wsfs.entry:
                if wsf.title.text == worksheetName_ :
                    #ok it is there, we just didn't know about it
                    self._addToMap(wsf)
                    break
            else:
                #no, it doesn't even exist
                if create_:
                    wsf = self.gd_client.AddWorksheet(worksheetName_, 1,1, self.key)
                    self._addToMap(wsf)
                else:
                    #sorry, it doesn't exist and we are not supposed to create it.
                    raise (Exception('Worksheet doesn\'t exist',worksheetName_))
        #ok, so by now we have the following cases:
        #doesn't exist, and don't create -> already dead
        #it didn't exist, but created and added to our database
        #existed, but had to add to database
        #existed and was in our database
        #in any case wsf is what we need it to be.
        
        
        cellfeed = self.gd_client.GetCellsFeed(self.key,self._WorksheetIDFromName(worksheetName_))
        if cellFeedOnly_:
            return cellfeed
            
        else:
                
            my_time = self.ws_map[worksheetName_][1]
            last_update_time = str2secs(wsf.updated.text)
    
            if last_update_time > my_time:
                # print('Reloading')
                #cellfeed = self.gd_client.GetCellsFeed(self.key,self._WorksheetIDFromName(worksheetName_))
                self.ws_map[worksheetName_][2] = cellFeedToDict(cellfeed)
                self.ws_map[worksheetName_][1] = last_update_time
                
            return self.ws_map[worksheetName_][2]
         
    def Read(self,worksheetName_,position_=(1,1),rows_=0,cols_=0, readString_ = False, allowEmpty_ = None):
        """
        
        Args:
          readString_: [bool] Whether to return strings (arbitrary length) or numbers in the array
          allowEmpty_: [anything] If not None: allow empty values in reading when both rows_ and cols_ are given.
              Its value is the default value. 
        """
        #at least one of rows_ or cols_ should be given (>0)
        
        if (rows_<0) or (cols_<0) or ((rows_+cols_) < 1):
            return (numpy.matrix(()))
            

        cellDict = self._cacheWS(worksheetName_,create_=False)            
        #Now we can start to do real work

        if (rows_ > 0) and (cols_ >0) and not (allowEmpty_ is None):
            readRows=rows_
            readCols=cols_
            
        else:
        #First we find out how many rows and columns we need to read if one of them is not given.
        #If both given we need to check if table holds the required data.
        #The number of rows/columns to read will be stored in variables readRows/readCols
            
            if rows_ > 0:
                readCols=sys.maxint
                readRows=rows_
                for r in range(position_[0],position_[0]+rows_):
                    if not cellDict.has_key(r):
                        raise (Exception('Row not found',r))
                    c = position_[1]
                    while isInDict(cellDict,r,c) and ((cols_ == 0) or (c<position_[1]+cols_)):
                        c += 1
                    if (cols_>0) and (c-position_[1] < cols_):
                        raise (Exception('Not enough columns in row',r))
                    readCols = min(readCols,c-position_[1])
                    if readCols == 0:
                        raise (Exception('Some rows have 0 columns'))
               
                        
            else: #rows is 0
                readCols = cols_
                r = position_[0]
                enough_cols = True
                while cellDict.has_key(r) and enough_cols:
                    c=position_[1]
                    while cellDict[r].has_key(c) and (c < position_[1]+cols_):
                        c += 1
                    if c<position_[1]+cols_:
                        enough_cols = False
                    else:
                    	r += 1
                
                #if r == position_[0]:
                    #raise(Exception('No rows with sufficient number of columns'))
                    
                readRows = r-position_[0]
                
                
        #ok, now we know what we need, only need to read        


        if readString_:
            a = numpy.empty([readRows,readCols],dtype=object)
        else:
            a = numpy.empty([readRows,readCols])

        if not allowEmpty_ is None:
            reader = lambda r,c : cellDict[r][c] if (cellDict.has_key(r) and cellDict[r].has_key(c)) else allowEmpty_
        else:
            reader = lambda r,c : cellDict[r][c]


        for r in xrange(position_[0],position_[0]+readRows):
            for c in xrange(position_[1],position_[1]+readCols):
                #a[r-position_[0]][c-position_[1]] = cellDict[r][c]
                a[r-position_[0]][c-position_[1]] = reader(r,c)
                
                
        return (a)
            

    def Store(self,matrix_,worksheetName_, (start_r_,start_c_)=(1,1)):
        """Stores a "matrix" in a google spreadsheet
        
        Args:
          matrix_: [anything which can be used to construct numpy.matrix]
            The matrix to store
            
          worksheetName_: [string] The name of the worksheet
          
          (start_r_,start_c_): [pair of int] Upper left corner of the storing position
        """
        myMatrix = numpy.matrix(matrix_)
        self._cacheWS(worksheetName_,create_=True)
        #ok, now it exists and is in our cache
        
        #check size; resize if necessary 
        
        #ws = self.gd_client.GetWorksheetsFeed(self.key,self._WorksheetIDFromName(worksheetName_))
        ws = self.ws_map[worksheetName_][3]
        ws_r = int(ws.row_count.text)
        ws_c = int(ws.col_count.text)
        
        matrix_r , matrix_c= myMatrix.shape
        needed_r = max(start_r_+matrix_r-1,ws_r)
        needed_c = max(start_c_+matrix_c-1,ws_c)
        
        if (needed_r > ws_r) or (needed_c > ws_c):
            ws.row_count.text = str(needed_r)
            ws.col_count.text = str(needed_c)
            self.gd_client.UpdateWorksheet(ws)
        
        #get the feed 
        query = gdata.spreadsheet.service.CellQuery()
        query.return_empty = "true" 
        query.min_row = str(start_r_)
        query.max_row = str(start_r_+matrix_r-1)
        query.min_col = str(start_c_)
        query.max_col = str(start_c_+matrix_c-1)
        cells = self.gd_client.GetCellsFeed(self.key, wksht_id=self._WorksheetIDFromName(worksheetName_), query=query)

        #create batch request
        batchRequest = gdata.spreadsheet.SpreadsheetsCellsFeed()
        #.add changes to batchrequest
        
        pos = 0
        for r in xrange(matrix_r):
            for c in xrange(matrix_c):
                cells.entry[pos].cell.inputValue = str(myMatrix[r,c])
                batchRequest.AddUpdate(cells.entry[pos])
                pos += 1
        
        updated = self.gd_client.ExecuteBatch(batchRequest, cells.GetBatchLink().href)
        
        
    def _KeyFromSpreadsheetName(self,spreadsheet_):
        f = self.gd_client.GetSpreadsheetsFeed()
        for fs in f.entry:
            if fs.title.text == spreadsheet_:
                ip=fs.id.text.split('/')
                return(ip[-1])
        #name not found:
            #need to create it
  
    def _WorksheetIDFromName(self,worksheetName_):
        return (self.ws_map[worksheetName_][0])
        
def isInDict(dict,r,c):
    try:
        dict[r][c]
    except:
        return False
    else:
        return True
    
def cellFeedToDict(cellFeed_, constructor_=lambda x:x):
    """
    Takes a gdata cellfeed and turns it into a dictionary of dictionary of floats,
    with row and column numbers as keys.
    """
    cell_dict=defaultdict(dict)
    for cf1 in cellFeed_.entry:
        row = int(cf1.cell.row)
        col = int(cf1.cell.col)
        
        try:
            value = constructor_(cf1.cell.inputValue)
        except ValueError:
            #we only store numbers. Needs less storage, and need to convert only once
            #of course we have to convert all cells, even if they will not be used.
            pass
        else:
            cell_dict[row][col] = value
    return cell_dict
    
def str2secs(str_):
    """
    >>> str2secs('2012-05-08T14:30:41.407Z')
    1336480241.0
    """
    date_time = map(int,str_.replace('-',':').replace('T',':').replace('.',':').split(':')[0:6])
    date_time.extend([0,0,1])
    st = time.struct_time(date_time)
    return time.mktime(st)
    
if __name__ == "__main__":
    import doctest
    doctest.testmod()
    
    import MatrixSpreadsheet_test as mst
    import unittest
    suite = unittest.TestLoader().loadTestsFromTestCase(mst.MatrixSpreadsheet_test)
    unittest.TextTestRunner(verbosity=2).run(suite)
