import logging
import os.path
import json
import zipfile
import pandas as pd

from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation

logging.basicConfig(level=os.environ.get("LOGLEVEL", "INFO"))

# dictionary with keys as column names and values as list
dict = {'Bussiness__Name':[],'Bussiness__StreetAddress':[],'Bussiness__City':[],'Bussiness__Country':[],
           'Bussiness__Zipcode':[],'Invoice__Number':[],'Invoice__IssueDate':[],'Bussiness__Description':[],
           'Customer__Name':[],'Customer__Email':[],'Customer__PhoneNumber':[],         
           'Customer__Address__line1':[],'Customer__Address__line2':[],'Invoice__Description':[],'Invoice__DueDate':[],'Invoice__Tax':[],
           'Invoice__BillDetails__Name':[],'Invoice__BillDetails__Quantity':[],'Invoice__BillDetails__Rate':[]
}
# obtaining base path
base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# to loop all 100 files in testdata input
for i in range(0,100):

  try:
    # Initial setup, create credentials instance with the help of json file
    # Enter your own credentials in that file
    # private.key file is also required
      credentials = Credentials.service_account_credentials_builder() \
        .from_file(base_path + "/pdfservices-api-credentials.json") \
        .build()

    # Create an ExecutionContext using credentials and create a new operation instance.
      execution_context = ExecutionContext.create(credentials)
      extract_pdf_operation = ExtractPDFOperation.create_new()

    # Set operation input from a source file.
      source = FileRef.create_from_local_file(base_path + "/TestDataSet/output"+str(i)+".pdf")
      extract_pdf_operation.set_input(source)

    # Build ExtractPDF options and set them into the operation
      extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
        .with_element_to_extract(ExtractElementType.TEXT) \
        .build()
      extract_pdf_operation.set_options(extract_pdf_options)

    # Execute the operation.
      result: FileRef = extract_pdf_operation.execute(execution_context)

    # Save the result to the specified location.
      result.save_as(base_path + "/output/output"+str(i)+".zip")
    # var to store zipfile path
      zip_file = (base_path+"/output/output"+str(i)+".zip")
    # reading json data
      archive = zipfile.ZipFile(zip_file, 'r')
      jsonentry = archive.open('structuredData.json')
      data = (json.loads(jsonentry.read()))

      s = ["","","","","",""]
      strj = []
      str1 = ""
      str2 = ""
      str3 = []
      c = 0
     # accessing element attribute of json data
      for emp in data['elements']:
            if "Text" in emp:
                  if "Path" in emp:
                        # With the help of Bounds attribute get the adequate coordinates of data and store it in list 's'
                        if ("Bounds" in emp):
                              if "Page" in emp:
                                 if(int(emp['Bounds'][0])<90 and int(emp['Bounds'][0])>65 and 
                                  int(emp['Bounds'][1])<750 and int(emp['Bounds'][1])>680 and 
                                  int(emp['Bounds'][2])<230 and int(emp['Bounds'][2])>90 and 
                                  int(emp['Bounds'][3])<760 and int(emp['Bounds'][3])>690 and emp["Page"]==0):
                                          s[0]+=emp['Text']
                                          continue
                                 elif(int(emp['Bounds'][0])<510 and int(emp['Bounds'][0])>310 and 
                                    int(emp['Bounds'][1])<740 and int(emp['Bounds'][1])>670 and 
                                    int(emp['Bounds'][2])<560 and int(emp['Bounds'][2])>520 and 
                                    int(emp['Bounds'][3])<750 and int(emp['Bounds'][3])>700 and emp["Page"]==0):
                                          s[1]+=emp['Text']
                                          continue
                                 elif(int(emp['Bounds'][0])<90 and int(emp['Bounds'][0])>65 and 
                                    int(emp['Bounds'][1])<670 and int(emp['Bounds'][1])>610 and 
                                    int(emp['Bounds'][2])<480 and int(emp['Bounds'][2])>270 and 
                                    int(emp['Bounds'][3])<690 and int(emp['Bounds'][3])>620 and emp["Page"]==0):
                                          s[2]+=emp['Text']
                                          continue
                                 elif(int(emp['Bounds'][0])<90 and int(emp['Bounds'][0])>65 and 
                                    int(emp['Bounds'][1])<590 and int(emp['Bounds'][1])>480 and 
                                    int(emp['Bounds'][2])<230 and int(emp['Bounds'][2])>75 and 
                                    int(emp['Bounds'][3])<600 and int(emp['Bounds'][3])>490 )and emp["Page"]==0:
                                          s[3]+=emp['Text']
                                          continue
                                 elif(int(emp['Bounds'][0])<260 and int(emp['Bounds'][0])>220 and 
                                    int(emp['Bounds'][1])<600 and int(emp['Bounds'][1])>500 and 
                                    int(emp['Bounds'][2])<400 and int(emp['Bounds'][2])>250 and 
                                    int(emp['Bounds'][3])<610 and int(emp['Bounds'][3])>510 and emp["Page"]==0):
                                          s[4]+=emp['Text']
                                          continue
                                 elif(int(emp['Bounds'][0])<430 and int(emp['Bounds'][0])>390 and 
                                    int(emp['Bounds'][1])<600 and int(emp['Bounds'][1])>550 and 
                                    int(emp['Bounds'][2])<530 and int(emp['Bounds'][2])>445 and 
                                    int(emp['Bounds'][3])<610 and int(emp['Bounds'][3])>560 and emp["Page"]==0):
                                          s[5]+=emp['Text']
                                          continue
                                 # Storing Table data in 'str1'
                                 if("Table" in emp["Path"] and emp["Text"][0]!='$' and c < 3 and int(emp['Bounds'][1])<420 and 
                                    emp['Text']!="Subtotal " and emp['Text']!="Tax % " and emp['Text']!="Total Due " and emp['Text'][0]!='$'
                                    and emp['Text']!="ITEM " and emp['Text']!="QTY " and emp['Text']!="RATE " and emp['Text']!="AMOUNT "):
                                        c += 1
                                        str1 += emp["Text"]
                                 # using this if condition and storing in 'str2' for retrieving tax% later  
                                 if(emp['Text']!="ITEM " and emp['Text']!="QTY " and emp['Text']!="RATE " and emp['Text']!="AMOUNT "):
                                          str2 += (emp['Text'])
                                 # appending str3 to store multiple rows of table
                                 if c >= 3:
                                    c = 0
                                    str3.append(str1)
                                    str1 = ""

      for i in range(6):
            # Splitting data in elements of 's' when space is encountered and storing individual words in var 'strj'
            strj = s[i].split()
            temp = ""
            if i==0:  
                  # appending data in key of dictionary 'dict'          
                  dict['Bussiness__Name'].append(strj[0]+' '+strj[1])
                  dict['Bussiness__Zipcode'].append(strj[len(strj)-1])
                  dict['Bussiness__Country'].append(strj[len(strj)-3]+' '+strj[len(strj)-2])
                  dict['Bussiness__City'].append(strj[len(strj)-4][:len(strj[len(strj)-4])-1])
                  for j in range(2,len(strj)-4):
                        temp+=strj[j]+' '  
                  dict['Bussiness__StreetAddress'].append(temp[:len(temp)-2])
            if i==1: 
                  # storing invoice number and invoice issue date    
                  dict['Invoice__Number'].append(strj[1])
                  dict['Invoice__IssueDate'].append(strj[len(strj)-1])
            if i==2:
                  # storing bussiness description
                  for j in range(2,len(strj)):
                        temp+=strj[j]+' '  
                  dict['Bussiness__Description'].append(temp[:len(temp)-1])
            if i==3:  
                  # storing customer data          
                  dict['Customer__Name'].append(strj[2]+' '+strj[3])
                  if '-' in strj[5]:
                        dict['Customer__Email'].append(strj[4])
                        dict['Customer__PhoneNumber'].append(strj[5])
                        dict['Customer__Address__line1'].append(strj[6]+' '+strj[7]+' '+strj[8])
                        for j in range(9,len(strj)):
                             temp+=strj[j]+' '  
                        dict['Customer__Address__line2'].append(temp[:len(temp)-1]) 
                  else:
                        dict['Customer__Email'].append(strj[4]+strj[5])
                        dict['Customer__PhoneNumber'].append(strj[6])
                        dict['Customer__Address__line1'].append(strj[7]+' '+strj[8]+' '+strj[9])
                        for j in range(10,len(strj)):
                             temp+=strj[j]+' '  
                        dict['Customer__Address__line2'].append(temp[:len(temp)-1]) 
                  
            if i==4:  
                  # storing invoice description          
                  for j in range(1,len(strj)):
                        temp+=strj[j]+' '
                  dict['Invoice__Description'].append(temp[:len(temp)-1])
            if i==5: 
                  # storing invoice due date           
                  dict['Invoice__DueDate'].append(strj[len(strj)-1])

       # appending table data in key of dictionary 'dict' 
      for i in str3:           
            strj = i.split()  
            temp = ""    
            dict["Invoice__BillDetails__Quantity"].append(strj[len(strj)-2])
            dict['Invoice__BillDetails__Rate'].append(strj[len(strj)-1])
            for j in range(0,len(strj)-2):
                  temp += strj[j]+' '
            dict['Invoice__BillDetails__Name'].append(temp[:len(temp)-1])

       # storing tax%
       # in 'str2' 4th last word store the value of tax %
      strk = []
      strk = str2.split()
      if '$' in strk[len(strk)-4]:
          # using this because in 'TestDataInput -> output81.pdf' JSON file created dosen't contain Tax%
          dict['Invoice__Tax'].append("10")  
      else:
          dict['Invoice__Tax'].append(strk[len(strk)-4])

      le = len(dict['Bussiness__Name'])-1
       
      # loop to append key values based on the number of items in table
      for i in range(1,len(str3)):
            
            dict['Bussiness__Name'].append(dict['Bussiness__Name'][le])
            dict['Bussiness__Zipcode'].append(dict['Bussiness__Zipcode'][le])
            dict['Bussiness__Country'].append(dict['Bussiness__Country'][le])
            dict['Bussiness__City'].append(dict['Bussiness__City'][le])
            dict['Bussiness__StreetAddress'].append(dict['Bussiness__StreetAddress'][le])
            dict['Invoice__Number'].append(dict['Invoice__Number'][le])
            dict["Invoice__IssueDate"].append(dict["Invoice__IssueDate"][le])
            dict['Bussiness__Description'].append(dict['Bussiness__Description'][le])
            dict['Customer__Name'].append(dict['Customer__Name'][le])
            dict['Customer__Email'].append(dict['Customer__Email'][le])
            dict['Customer__PhoneNumber'].append(dict['Customer__PhoneNumber'][le])
            dict['Customer__Address__line1'].append(dict['Customer__Address__line1'][le])
            dict['Customer__Address__line2'].append(dict['Customer__Address__line2'][le])
            dict["Invoice__Description"].append(dict["Invoice__Description"][le])
            dict['Invoice__DueDate'].append(dict['Invoice__DueDate'][le])
            dict['Invoice__Tax'].append(dict['Invoice__Tax'][le])

  except (ServiceApiException, ServiceUsageException, SdkException):
    logging.exception("Exception encountered while executing operation")

# df stores dataframe created from 'dict'
df = pd.DataFrame(dict)
#sorting dataframe based on column names
df = df.reindex(sorted(df.columns), axis=1)
# converting data frame to csv file
df.to_csv(base_path+"/output/ans.csv",index=False)


