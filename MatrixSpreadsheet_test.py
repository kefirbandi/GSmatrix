# -*- coding: utf-8 -*-
"""
Created on Fri May 11 12:43:42 2012

@author: agefferth
"""

import MatrixSpreadsheet as ms
import random
import numpy
import unittest

USER = '...@gmail.com'
PASS = '...'

class MatrixSpreadsheet_test(unittest.TestCase):

    def setUp(self):
        self.m = ms.MatrixSpreadsheet(USER, PASS, spreadsheetName_ = 'Test 1') #Supply existing spreadsheet name
        self.wsname = 'TEST001'+str(random.randint(0,1e9))
        
        
    def test_open(self):
        #Trying to open with both name and key supplied
        self.assertRaises(Exception,ms.MatrixSpreadsheet,USER, PASS, 'Test 1','1234566')

        #Opening with key
        ms.MatrixSpreadsheet(USER, PASS, spreadsheetKey_ = '0Aov...2JOcHc') #Supply existing spreadsheet key

    def test_cacheNotFound(self):
        self.assertRaises(Exception,self.m._cacheWS,self.wsname)
        
    def test_readWrite(self):
        data =  numpy.array([range(4),range(4,8)])
        
        #Store in a new database
        self.m.Store(data,self.wsname)

        #Read with number of ROWS specified
        newData = self.m.Read(self.wsname,rows_=numpy.shape(data)[0])
        self.assertTrue((data == newData).all())
        
        #Read with number of COLUMNS specified
        newData = self.m.Read(self.wsname,cols_=numpy.shape(data)[1])
        self.assertTrue((data == newData).all())
        
        #Read with BOTH number of ROWS amd COLUMNS specified
        newData = self.m.Read(self.wsname,rows_=numpy.shape(data)[0],cols_=numpy.shape(data)[1])
        self.assertTrue((data == newData).all())
   
        #Store in existing database
        self.m.Store(data+1,self.wsname)
        
        #Read with number of COLUMNS specified
        newData = self.m.Read(self.wsname,cols_=numpy.shape(data)[1])
        self.assertTrue(((data+1) == newData).all())
   
   
        #Remove temporary worksheet
        wsf = self.m.gd_client.GetWorksheetsFeed(self.m.key,self.m._WorksheetIDFromName(self.wsname))
        self.m.gd_client.DeleteWorksheet(wsf)

if __name__ == '__main__':
    unittest.main()
