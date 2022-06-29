#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Feb  3 15:44:15 2022

@author: phykc
"""
from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd
import h_codes_finder as hcf

class Coshh:
    """Find the COSHH form fields from template and upload the cossh data from excel"""
    def __init__(self, wordtemplate='COSHH form.docx', exceltemplate='COSHHdata.xlsx'):
        self.wordtemplate=wordtemplate
        self.exceltemplate=exceltemplate
        self.document= MailMerge(self.wordtemplate)
        self.df=pd.read_excel('COSHHdata.xlsx')
        self.dfcodes=pd.read_excel('hazardcodes.xlsx')
        self.checkdf=self.df.isnull()
        self.checkdf1=self.dfcodes.isnull()
    def create(self):
        name=self.df.loc[0,'FullName']
        title=self.df.loc[0,'Procedure name']
        # Insert the basic information intot he fields
        # print(self.document.get_merge_fields())
        # print(self.df.iloc[1,12])
        checklist=[]
        for r in range(13):
            # print(self.df.iloc[r,11])
            if self.df.iloc[r,11]=='y':
                checklist.append('\u2612')
            else:
                checklist.append('\u2610')
        thewindex=int(self.df.iloc[r+1,11])-1
        
        options=['\u2610','\u2610','\u2610','\u2610','\u2610']
        for b in range(5):
            if b ==thewindex:
                options[b]='\u2612'
        
        self.document.merge(Safety_Specs=checklist[0],Hood=checklist[1],Labcoat=checklist[2],Face_shield=checklist[3], Dust_mask=checklist[4],Thin_gloves=checklist[5], Thick_gloves=checklist[6], Assesor=name,CO2_for_fire=checklist[7],Foam_for_fire=checklist[8],Contain_spill_with_absorbent=checklist[9],Soda_ash=checklist[10],Wash_with_water=checklist[11],Seek_medical=checklist[12], Authoriser=self.df.loc[0,'Authoriser'], Supervisor=self.df.loc[0,'Supervisor'], Procedure=self.df.loc[0,'Procedure'], Date='{:%d-%b-%Y}'.format(date.today()), FullName=name,Procedure_name=title, Source=self.df.loc[0,'Source'], Status=self.df.loc[0,'Status'], Work_place=self.df.loc[0,'Work Place'], o1=options[0],o2=options[1],o3=options[2],o4=options[3],o5=options[4])
        # To populate a table correctly we need to collect substances and the harards from the list of codes.
        substances=[]
        hazards=[]
        # First find the list of substances
        for n in range(len(self.df)):
            if self.checkdf.loc[n,'Substances']==False:
                substances.append(self.df.loc[n,'Substances'])
        # Next find each hazard code as a list.
        for i,subs in enumerate(substances):
            hc=hcf.Hazard_codes(subs)
            hzstr=hc.findthem()
            hazardstr=hzstr
            hazstr=hazardstr.strip()
            haz=hazstr.replace(' ','')
            if ',' in haz:
                hazlist=haz.split(',')
            elif '-' in haz:
                hazlist=haz.split('-')
            else:
                hazlist=[haz]
            # print(hazlist)    
            haztextlist=[]
            # Collect the text for each code and place them in a formated string. After each hazard code a new line is created.
            for code in hazlist:
                indx=self.dfcodes[self.dfcodes['Hazard Code']==code].index.values
                if len(indx)!=0:
                    statement=self.dfcodes.loc[indx[0],'Hazard Statement']
                    if self.checkdf1.loc[indx[0],'Addition']==False:
                        statement+=' '+self.df.loc[indx[0],'Addition']
                    haztextlist.append(statement+'.'+'\n') 
                else:
                    print('Hazard Code(s) not found. Complete this part yourself.')
            
            harazrdstring=''.join(haztextlist)
            hazards.append(harazrdstring[:-2])
        # Place the substances and corresponding hazard text (str) in to a list of dicts.
        chemicalhazards=[]             
        for n in range(len(substances)):
            chemicalhazards.append({'Substances':substances[n], 'Hazards':hazards[n]})
        # Populate the table
        self.document.merge_rows('Substances', chemicalhazards)
        # Save the COSHH form using the title name
        self.document.write(title+'_COSHH.docx')

if __name__=='__main__':
    # Create a coshh instance using the .docx and .xlsx templates
    cosshform=Coshh(wordtemplate='COSHH form_temp.docx', exceltemplate='COSHHdata.xlsx')
    cosshform.create()
    print('Created a COSHH .docx file')