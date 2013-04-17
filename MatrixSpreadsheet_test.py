# -*- coding: utf-8 -*-
"""
Created on Fri May 11 12:43:42 2012

@author: agefferth
"""

import MatrixSpreadsheet as ms
import random
import numpy
import unittest

reload(ms)
class MatrixSpreadsheet_test(unittest.TestCase):

    def setUp(self):
        self.m = ms.MatrixSpreadsheet('your_email','your_password','Test 1')
        self.wsname = 'TEST001'+str(random.randint(0,1e9))
        
        
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
