import json
import pandas
from shareplum import Site
from shareplum import Office365
import time
import pandas as pd
from datetime import date
from json import JSONEncoder
import yagmail

import os
from decouple import config

def remove_umlaut(string):
    """
    Removes umlauts from strings and replaces them with the letter+e convention
    :param string: string to remove umlauts from
    :return: unumlauted string
    """
    u = 'ü'.encode()
    U = 'Ü'.encode()
    a = 'ä'.encode()
    A = 'Ä'.encode()
    o = 'ö'.encode()
    O = 'Ö'.encode()
    ss = 'ß'.encode()

    string = string.encode()
    string = string.replace(u, b'ue')
    string = string.replace(U, b'Ue')
    string = string.replace(a, b'ae')
    string = string.replace(A, b'Ae')
    string = string.replace(o, b'oe')
    string = string.replace(O, b'Oe')
    string = string.replace(ss, b'ss')

    string = string.decode('utf-8')
    return string
 
class Request(object):
    def __init__(self, name, item ):
        self.name = name
        self.item = item
        self.email = ""



user = config('iptvmail')

pwd = config('pwd')




emailist = []
authcookie = Office365('https://reutlingenuniversityde.sharepoint.com/', username= config('username') , password=config('password')).GetCookies()
site = Site('https://reutlingenuniversityde.sharepoint.com/sites/Test01', authcookie=authcookie)

sp_list = site.List('KK_Test_List_01')
sp_list1 = site.List('Taschen_Liste')
fields = ['Person', 'Item']
fields1 =['Person','Tasche']
query = {'Where': [('Lt', 'EndDate', str(date.today()))]}
                          

data = pd.DataFrame(sp_list.GetListItems(fields=fields, query=query)[0:]).dropna(thresh=1)


data1 = pd.DataFrame(sp_list1.GetListItems(fields=fields1,query=query)[0:]).dropna(thresh=1)

data.reset_index(drop = True, inplace = True)
data1.reset_index(drop = True, inplace = True)

dpd= data.append(data1.rename(columns = {"Tasche":"Item"}), ignore_index=True).dropna(thresh=1)
try:
    dpd = dpd[dpd.Person.notnull()]
except: 
    print('Niemand wurde gefunden')
    quit()

for index,row in dpd.iterrows():

        emailist.append(Request(row['Person'], row['Item']))

for request in emailist:

    request.name
    s = request.name.split()
    sl = s.pop()
    email= lambda request, s: '_'.join(s)+"."+sl+"@cs.Reutlingen-University.DE"
    
    request.email = remove_umlaut(email(request,s))
    print(request.email)
    
    print('succes')
    try:
        yag = yagmail.SMTP(user, pwd)
        yag.send(request.email, subject="Rückgabe "+request.item, contents="Hallo "+s[0] + "," + "\n" + "Das Item "+ request.item +" ist bereits länger ausgeliehen als in der App eingetragen"+ "\n"+ "Vielen Dank" +"\n" + "dein IPtv Admin" )
    except:
        print('Email existiert nicht')
        
        
       


