#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 20 19:33:49 2022

@author: phykc
"""
import pandas as pd # library for data analysis
import requests # library to handle requests
from bs4 import BeautifulSoup # library to parse HTML documents
# get the response in the form of html
class Hazard_codes:
    def __init__(self, chemname):
        self.chemname=chemname.replace(' ','_')
        self.searchname='Hazard statements'
    def findthem(self):
        self.wikiurl="https://en.wikipedia.org/wiki/"+self.chemname
        self.response=requests.get(self.wikiurl)
        if self.response.status_code==200:
            # parse data from the html into a beautifulsoup object
            soup = BeautifulSoup(self.response.text, 'html.parser')
            self.chemtable=soup.find('table',{'class':"infobox ib-chembox"})
            df=pd.read_html(str(self.chemtable))
            # convert list to dataframe
            df=pd.DataFrame(df[0])
            listnames=df.iloc[:,0].to_list()
            try:
                ival=listnames.index(self.searchname)
                codes=df.iloc[ival,1]
                if codes[-1]==']':
                    self.codes=codes[:-3]
                else:
                    self.codes=codes
                return self.codes
            except ValueError:
                return ""
            
        else:
            return ""

if __name__=="__main__":        
    hc=Hazard_codes('hydrochloric acid')
    hazardcodes=hc.findthem()
    print(hazardcodes)
