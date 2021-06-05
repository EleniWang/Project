'''
The python crawler realizes the automatic update of the price list. The specific operation is to read the links of the price list one by one, take a screenshot with PhantomJS, then paste it into the PDF, and then call the Dropbox API to update the price. 
'''
import csv
import xlrd
import os
import time
import sys
import glob
import fitz
import shutil
import re
import configparser
import dropbox
from reportlab.lib.pagesizes import portrait
from reportlab.pdfgen import canvas
from PIL import Image
from selenium import webdriver

# read in the excel to get the link
wb = xlrd.open_workbook('Desktop' + '\\' + 'pricelist.xlsx')
sh = wb.sheet_by_name('worksheet1')
n = int((sh.nrows+1)/3)


directory_time = time.strftime("%Y-%m-%d", time.localtime(time.time()))
try: 
    File_Path = os.getcwd() + '\\' + 'Desktop' + '\\' + directory_time + '\\'
    if not os.path.exists(File_Path):
        os.makedirs(File_Path)
        print("New directory created：%s" % File_Path)
    else:
        print("Directory exist")
except BaseException as msg:
    print("Failed to create new directory：%s" % msg)

# screen shot and save in temporary folder 
driver1 = webdriver.PhantomJS()
for i in range(n):
    driver1.get(sh.cell(i*3+1,2).value)
    driver1.save_screenshot(os.getcwd() + '\\' 'Desktop' + '\\' + directory_time +  '\\' + str(i) + '.png')
    
# save as PDF
for i in range(n):
    a = sh.cell(i*3,5).value
    ImgFile = Image.open(os.getcwd() + '\\' + 'Desktop' + '\\' + directory_time +  '\\' + str(i) + '.png')
    if ImgFile.mode == 'RGBA':
        ImgFile = ImgFile.convert("RGB")
    ImgFile.save(os.getcwd() + '\\' + 'Desktop' + '\\' + directory_time +  '\\' + str(a) + '_' + directory_time + '.pdf',"PDF")
    ImgFile.close()

# delete temporary file
def del_files(path):
    for root , dirs, files in os.walk(path):
        for name in files:
            if name.endswith(".png"):
                os.remove(os.path.join(root, name))
                print ("Delete File: " + os.path.join(root, name))
 
if __name__ == "__main__":
    path = os.getcwd() + '\\' + 'Desktop' + '\\' + directory_time
    del_files(path)


# call the Dropbox API
class TransferData:
    def __init__(self, access_token):
        self.access_token = access_token

    def upload_file(self, file_from, file_to):
        dbx = dropbox.Dropbox(self.access_token)

        with open(file_from, 'rb') as f:
            dbx.files_upload(f.read(), file_to)

def main():
    access_token = 'xxx'
    transferData = TransferData(access_token)
    
    for i in range(n):
        a = sh.cell(i*3,5).value
        file_from = os.getcwd() + '\\' + 'Desktop' + '\\' + directory_time + '\\' + str(a) + '_' + directory_time + '.pdf' # This is name of the file to be uploaded
        file_to =sh.cell(i*3,6).value + str(a) + '_' + directory_time + '.pdf'  # This is the full path to upload the file to, including name that you wish the file to be called once uploaded.
        print(transferData.upload_file(file_from, file_to))
if __name__ == '__main__':
    main()



# delete all the file no longer useful
import os
import shutil
filelist=[]
rootdir=os.getcwd() + '\\' + 'Desktop' + '\\' + directory_time                       #Take the path of the deleted folder, and the final result deletes the img folder 
filelist=os.listdir(rootdir)                #List all file names in this directory 
for f in filelist:
    filepath = os.path.join( rootdir, f )   #Map file name to absolute path 
    if os.path.isfile(filepath):            #Determine whether the file is a file or folder 
        os.remove(filepath)                 #If it is a file, delete it directly 
        print(str(filepath)+" removed!")
    elif os.path.isdir(filepath):
        shutil.rmtree(filepath,True)        #If it is a folder, delete the folder and all files in the folder 
        print("dir "+str(filepath)+" removed!")
shutil.rmtree(rootdir,True)                 #Finally delete the img total folder 
print("successfully deleted ")

path = os.getcwd() + '\\' + 'Desktop' + '\\' + 'Bathla portal passcode .xlsx'  # directory path
if os.path.exists(path):  # if exist
    os.remove(path)  
    #os.unlink(path)
else:
    print('no such file:%s'%my_file)  






