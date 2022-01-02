import PyPDF2
import pandas as pd
import os
import math
import io
import glob
from openpyxl import load_workbook
from shutil import copyfile
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import sys, getopt
import datetime
import tabula as tb
import tabulate
import win32api
from win32com import client
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from PyPDF4.pdf import PdfFileReader, PdfFileWriter
import win32com.client
from PyPDF2 import PdfFileReader, PdfFileWriter
from getpass import getpass
import PyPDF2
import io
import time
import glob



#File Search Related Tools
    
def search_drives(extension):
    
    """
    This function is used to all the files of a particular extension in your entire system.
    
    It accepts just one parameter 
    
    extension: the second mandatory parameter is the extension . if you want to search pdf file, you can mention it as "pdf" or ".pdf"
    
    
    This function returns a datafrae with the File Names and the Complete FIlepath
    
    
    
    """
  


    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]

    list_a=[]
    list_b=[]

    for i in drives:


        dir_path = os.path.dirname(os.path.realpath(i))



        for root, dirs, files in os.walk(dir_path):
            for file in files: 

                if file.endswith(extension):
                    search=root+'\\'+str(file)


                    list_a.append(search)
                    
                    sizes=os.path.getsize(search)/1024
                    
                    list_b.append(round(sizes,2))
                    
                    
    
    total_files = len(list_a)
    dict_a={"Complete File Path":list_a, "Size of the File (in Kb)":list_b }
    

    print( total_files,"- Search results")

    df = pd.DataFrame(dict_a)
    df = df.reset_index()
    df = df.rename(columns={"index":"S.N"})
    df['S.N'] = df.index + 1
    
    df["File Name"]=df["Complete File Path"].apply(lambda x:x.split("\\")[-1])
    
    folder=os.getcwd()
    
    df.to_excel(folder+"\Drive Search Results .xlsx",index=False)
    
    print(f"Complete List of {extension} Files with their Filepath exported in Excel. Location:  {folder}")
    return(df)


def search_folder(folder,extension):
    
    a=folder

    dir_path = os.path.dirname(os.path.realpath(a))
    list_a=[]
    list_b=[]
    

    for root, dirs, files in os.walk(dir_path):
        for file in files: 

            if file.endswith(extension):
                
                search=root+'\\'+str(file)
                

                list_a.append(search)
                
                sizes=os.path.getsize(search)/1024
                
                list_b.append(sizes)
                
    total_files = len(list_a)
    dict_a={"Complete File Path":list_a , "Size of the File (in Kb)":list_b}

    print( total_files,"- Search results")

    df = pd.DataFrame(dict_a)
    df = df.reset_index()
    df = df.rename(columns={"index":"S.N"})
    df['S.N'] = df.index + 1
    
    df["File Name"]=df["Complete File Path"].apply(lambda x:x.split("\\")[-1])
    
    
    df.to_excel(folder+"\Folder Search Results .xlsx",index=False)
    print(f"Complete List of {extension} Files with their Filepath exported in Excel. Location:  {folder}")
    return(df)




#Excel Related Tools


def split_excel(filepath ,cols , sheet=0, mode="file"):
    
    df = pd.read_excel(filepath,sheet_name=sheet)
    
    
    if mode=="file":

        df = pd.read_excel(filepath,sheet_name=sheet)
        pth = os.path.dirname(filepath)
        colslist = list(set(df[cols].values))



        for i in colslist:
            df[df[cols] == i].to_excel("{}/{}.xlsx".format(pth, i), sheet_name=i[0:25:1], index=False)

        print('Your data has been split into {} and {} files has been created.Click OK. \n All Files stored in same folder{}'.format(
                                ', '.join(colslist), len(colslist),pth))

        print("The names of the files are same as the name of the column items")
        return


    elif mode=="sheets":
        extension = os.path.splitext(filepath)[1]
        filename = os.path.splitext(filepath)[0]
        pth = os.path.dirname(filepath)
        newfile = os.path.join(pth, filename + '_Sheet_Split_Auto' + extension)
        df = pd.read_excel(filepath,sheet_name=sheet)
        colslist = list(set(df[cols].values))

        copyfile(filepath, newfile)
        for j in colslist:
            writer = pd.ExcelWriter(newfile, engine='openpyxl')
            for myname in colslist:
                mydf = df.loc[df[cols] == myname]
                mydf.to_excel(writer, sheet_name=str(myname[0:25:1]), index=False)
            writer.save()

        print('Your data has been split into {} and {} sheets has been created \n \n. File with all these sheets stored in  {new}.\n .'.format(', '.join(colslist), len(colslist),new=newfile))

        return



def combine_excel(folder,mode="file",sheet=0):


    files = glob.glob(os.path.join(folder, '*.xls*'))
    newfile = os.path.join(folder, 'All_ExcelFiles_Combined_Auto.xlsx')
    df = pd.DataFrame()
    
    if mode=="file":

        for f in files:

            data = pd.read_excel(f,sheet_name=sheet)
            data["File_Name"] = f
            df = df.append(data)


        df.to_excel(newfile, sheet_name='Combined', index=False)
        
        print(f"All the files in the folder {folder} have been combined into a single Excel File. \n\n The Combined Excel File stored in {newfile}")
        
        
    elif mode=="sheets":
        
        #in case we need to merge different sheets, we need to specify the complete file path and not only the directory
        
        filepath=folder


        pth = os.path.dirname(filepath)

        df = pd.DataFrame()

        df2 = pd.DataFrame()

        xl = pd.ExcelFile(filepath)


        newfile = os.path.join(pth, 'All_Sheets_Combined_Auto.xlsx')



        res = len(xl.sheet_names)


        while res>0:
            res-=1
            df=pd.read_excel(filepath,sheet_name=res)
            df2=df2.append(df)

        df2.to_excel(newfile, sheet_name='Combined', index=False)
        
        print(f"All the sheets in the  Excel file {filepath} have been combined into a single sheet named 'Combined' . \n \n  The New Excel File is stored in {newfile}")

            
    else:
        pass
        
    

def combine_txt(folder):
    
    import pandas as pd
    import glob
    import os

#     pth = os.path.dirname(filepath)
#     extension = os.path.splitext(filepath)[1]

    newfile = os.path.join(folder, 'Combined_Text_File_Auto.txt')

    filenames = glob.glob(folder + "/*.txt")

    df2 = pd.DataFrame()

    for files in filenames:
        df = pd.read_csv(files, sep="\t", low_memory=False,encoding='cp1252')

        df2 = df2.append(df)

    df2.to_csv(newfile, sep="\t")

    print(f'All text files in the selected folder have been merged and stored in {newfile}')




    
def combine_pdf(folder):


    filenames = glob.glob(folder + "/*.pdf")

    merged = PdfFileMerger()

    for files in filenames:
        merged.append(files)

    filename = os.path.splitext(folder)[0]


    newfile = os.path.join(folder, 'Combined_Pdf_File_Auto' + ".pdf")

    merged.write(newfile)
    merged.close()

    print('Output', 'All pdf files in the selected folder have been merged.\n Click on OK')
    
    
    #done




def split_pdf(filepath,type="page_wise",set=1):

    f = open(filepath, 'rb')
    pdf = PdfFileReader(f)

    if type=="page_wise":
                   

        
        for i in range(pdf.getNumPages()):
            
            writer= PdfFileWriter()  
              
            writer.addPage(pdf.getPage(i))
            
            extension = os.path.splitext(filepath)[1]
            
            pth = os.path.dirname(filepath)
            
            
            folder=os.path.dirname(filepath)
            newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Page_"+str(i+1)+"."+filepath.split("\\")[-1].split(".")[1]
        
            
            output = open(newfile, "wb")

            writer.write(output)

            output.close()
            
            print(f"The File has been saved as {newfile}")
            
        f.close()
        #done



    elif type=="cummulative":
        
        writer = PdfFileWriter()

        for i in range(0, pdf.getNumPages()):

            page = pdf.getPage(i)
            
            
            writer.addPage(page)
            
            
            extension = os.path.splitext(filepath)[1]
            
            pth = os.path.dirname(filepath)
            
#             newfile = os.path.join(pth, 'Page_1-' + str(i+1) + extension)
            
            folder=os.path.dirname(filepath)
            newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Page_1-"+str(i+1)+"."+filepath.split("\\")[-1].split(".")[1]
        
            
            output = open(newfile, "wb")

            writer.write(output)

            output.close()
            
            
            print(f"The File has been saved as {newfile}")
            
        f.close()

            

    elif type == "oddeven":
        writer1 = PdfFileWriter()
        writer2 = PdfFileWriter()

        for i in range(0, pdf.getNumPages()):

            if (i+1)%2 == 0:
                
                page = pdf.getPage(i)
                writer1.addPage(page)
            else:
                page = pdf.getPage(i)
                writer2.addPage(page)

        extension = os.path.splitext(filepath)[1]
        
        pth = os.path.dirname(filepath)
        
        folder=os.path.dirname(filepath)
        newfile_even=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Even Pages."+filepath.split("\\")[-1].split(".")[1]
        
        newfile_odd=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"Odd Pages."+filepath.split("\\")[-1].split(".")[1]
        
        
        

        output_even = open(newfile_even, "wb")
        output_odd = open(newfile_odd, "wb")

        writer1.write(output_even)

        writer2.write(output_odd)

        output_even.close()

        output_odd.close()
        f.close()

        print(f"The File has been saved as {newfile_even} & {newfile_odd}")
        #done


    elif type == "ranges":

        tot_page=pdf.getNumPages()
        files=math.ceil(tot_page/set)

        a=0

        for i in range(files):

            writer1 = PdfFileWriter()

            try:
                for j in range(0,set):

                    page = pdf.getPage(a)
                    a = a + 1

                    writer1.addPage(page)






                extension = os.path.splitext(filepath)[1]
                pth = os.path.dirname(filepath)
                
                
#               
                folder=os.path.dirname(filepath)
                newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Set_"+str(i+1)+"."+filepath.split("\\")[-1].split(".")[1]
        



                output = open(newfile, "wb")
                writer1.write(output)

                output.close()

            except IndexError:
                for j in range(0, tot_page - a):
                    page = pdf.getPage(a)
                    a = a + 1

                    writer1.addPage(page)





                extension = os.path.splitext(filepath)[1]
                pth = os.path.dirname(filepath)
                
                folder=os.path.dirname(filepath)
                newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Set_"+str(i+1)+"."+filepath.split("\\")[-1].split(".")[1]
        



                output = open(newfile, "wb")
                writer1.write(output)

                output.close()

            print(f"The File has been saved as {newfile}")


        f.close()


    elif type == "split_equal":

        tot_page = pdf.getNumPages()
        set_new=math.floor(tot_page/set)

        files = math.ceil(tot_page / set)

        a = 0

        for i in range(set):

            writer1 = PdfFileWriter()

            try:
                for j in range(0, set_new):
                    page = pdf.getPage(a)
                    a = a + 1

                    writer1.addPage(page)

                extension = os.path.splitext(filepath)[1]
                pth = os.path.dirname(filepath)
                
                folder=os.path.dirname(filepath)
                newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Set_"+str(i+1)+"."+filepath.split("\\")[-1].split(".")[1]
        

                output = open(newfile, "wb")
                writer1.write(output)

                output.close()

            except IndexError:
                for j in range(0, tot_page - a):
                    page = pdf.getPage(a)
                    a = a + 1

                    writer1.addPage(page)

                extension = os.path.splitext(filepath)[1]
                pth = os.path.dirname(filepath)
                
                folder=os.path.dirname(filepath)
                newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_Set_"+str(i+1)+"."+filepath.split("\\")[-1].split(".")[1]
        

                output = open(newfile, "wb")
                writer1.write(output)

                output.close()

            print(f"The File has been saved as {newfile}")

        f.close()



  



def sort_pdf(filepath):
    
    
    """
    This function is used to sort the pdf files in reverse order.
    
    Only One parameter is required i.e the Complete path to the file which is to be sorted.
    
    
    Output will be stoed in the sae folder as the original pdf file.
    
    The name of the pdf file will remain same, only the _sorted will be added at the last
    
    """
    

    output_pdf = PdfFileWriter()
    original_file=filepath.split("\\")[-1]

    with open(filepath, 'rb') as readfile:
        input_pdf = PdfFileReader(readfile)

        for page in reversed(input_pdf.pages):
            output_pdf.addPage(page)
            
        folder=os.path.dirname(filepath)
        filename=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_sorted."+filepath.split("\\")[-1].split(".")[1]
        with open(filename, "wb") as writefile:
            output_pdf.write(writefile)
            
        print(f"The pdf File{original_file} has been sored in reverse order and stored in same folder as original pdf file")
        
        return (output_pdf)
        


    

def encrypt_pdf(filepath,password="Password@1",type="function"):
    
    
    
    """
    
    This function is for setting the passwordin any pdf file.
    
    There are two ways of setting the password.
    
    We can provide a complete filepath and also provide a password we want to set.
    
    Input Parameters are as below:
    
    filepath= The complete path to the file needs to be provided. 
    
    password= This is an optional Parameter. By Default the password will be set as "Password@1".
                 It is highly recommended not to use the default password and set your own calue insde the password
                 
    
    type= Default value is function. This need not be changes unless you want the user to provide the password
    
    If user will provide the passwor, then select type as "user". This will prompt a input box 
    
    And user will have to enter the password in the input box
    
    Output will be stoed in the sae folder as the original pdf file.
    
    The name of the pdf file will remain same, only the _encrypted will be added at the last
    
    """
    
    

    
    if type=="function":
        
        pdfFile = open(filepath, 'rb')

        original_file=filepath.split("\\")[-1]

        # Create reader and writer object

        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        pdfWriter = PyPDF2.PdfFileWriter()
        
        # Add all pages to writer (accepted answer results into blank pages)
        
        for pageNum in range(pdfReader.numPages):
            pdfWriter.addPage(pdfReader.getPage(pageNum))
        # Encrypt with your password
        
        pdfWriter.encrypt(password)
        folder=os.path.dirname(filepath)
        filename=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_encrypted."+filepath.split("\\")[-1].split(".")[1]

        # Write it to an output file. (you can delete unencrypted version now)

        resultPdf = open(filename, 'wb')
        pdfWriter.write(resultPdf)
        resultPdf.close()

        print(f"The file {original_file} has been encrypted and stored in same folder ")
        return(resultPdf)
    
    elif type=="user":
        
        original_file=filepath.split("\\")[-1]

        print(f"The file {original_file} has been selected. What password do you want to set? ")
        file_pass=getpass("Enter the Password for this pdf File")
        
        encrypt_pdf(filepath,password=file_pass,type="function")
        
        print("The Password has been set")
        
    
    else:
        print("Some Error in Inputs given. Please Check Documentation")
    
    
        



def decrypt_pdf(filepath, password="Password@1" , type="function"):


    
    """

    This function is for decrypting the PDf files.


    The function takes 3 input parameters:

    filepath: This takes the complete path to the pdf file which is to be decrypted

    password: This will take the password which is used to open the psd file. The password can be passed as a function parameter .

    type: It has two options to be given as values.


    If the password is same for multiple files, then password can be passed as a functional parameter and same password will be used for each file.

    In this case, type to be selected as type="function"


    if we want the user to input the password each time for each pdf file , the the type needs to be selected as type="user"

    In this case, user will be asked for entering the password for decry

    





    """
    import os
    from PyPDF2 import PdfFileReader
    from PyPDF2 import PdfFileWriter
    
    
    if type=="function":

        folder = os.path.dirname(filepath)

        original_file=filepath.split("\\")[-1]


        output_path=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_decrypted."+filepath.split("\\")[-1].split(".")[1]


        with open(filepath, 'rb') as input_file,open(output_path, 'wb') as output_file:

            reader = PdfFileReader(input_file)
            reader.decrypt(password)

            writer = PdfFileWriter()

            for i in range(reader.getNumPages()):
            
                writer.addPage(reader.getPage(i))

            writer.write(output_file)

        print(f"The file {original_file} has been decrypted")
        
    
    elif type=="user":
        
        original_file=filepath.split("\\")[-1]

        print(f"The file {original_file} has been selected. Enter the password to decrypt ")
        
        time.sleep(1)
        
        file_pass=getpass("Enter the Password for this pdf File")
        
        decrypt_pdf(filepath,password=file_pass,type="function")
        
        print("The file has been decrypted \n \n")
        
    
    else:
        print("Some Error in Inputs given. Please Check Documentation")
        




def create_pdfpage(num, tmp):


    """

    This is an supporting function for the add_pagenum function.

    """
    c = canvas.Canvas(tmp)
    for i in range(1, num + 1):
        c.drawString((210 // 4) * mm, (4) * mm, str(i))
        c.showPage()
    c.save()





def addnumber_pdf(filepath):
    
    """
    This function is used for adding the page Number at the bottom of the pdf file.
    
    We only need to provde one parameter , i.e the complete file path to the pdf file
    
    The function will return a separate pdf with suffix as _numbers  and this pdf file will be stored in the same location.
    
    """

    from PyPDF2 import PdfFileWriter



    tmp = "__tmp.pdf"

    original_file=filepath.split("\\")[-1]

    output = PdfFileWriter()
    with open(filepath, 'rb') as f:
        pdf = PdfFileReader(f, strict=False)
        n = pdf.getNumPages()

        # create new PDF with page numbers
        create_pdfpage(n, tmp)

        with open(tmp, 'rb') as ftmp:
            numberPdf = PdfFileReader(ftmp)
            # iterarte pages
            for p in range(n):
                page = pdf.getPage(p)
                numberLayer = numberPdf.getPage(p)
                # merge number page with actual page
                page.mergePage(numberLayer)
                output.addPage(page)

            # write result
            if output.getNumPages():
                folder=os.path.dirname(filepath)
                new_file=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_numbered."+filepath.split("\\")[-1].split(".")[1]
        
                # newpath = filepath[:-4] + "_numbered.pdf"
                with open(new_file, 'wb') as f:
                    output.write(f)
        os.remove(tmp)
        
        print(f"The Page Number has been added in {original_file} and this numbered pdf has been saved as {new_file}")
  



def delete_pdfpage(filepath):
    
    original_file=filepath.split("\\")[-1]
    
    print(f"The pdf file {original_file} has been selected")
    
    time.sleep(1)

    n=int(input(" Enter no.of pages to delete : "))
    pages_to_delete = list(map(int,input("\nEnter page numbers(with spaces no commas) : ").strip().split()))[:n]
    
    
    print(pages_to_delete)

    #the below loop is added to substrct one from each item in th list
    
    #it is done because python starts counting from Zero, but human counting begins from one
    
    for i in range(len(pages_to_delete)):
        pages_to_delete[i] = pages_to_delete[i] - 1


    infile = PdfFileReader(filepath, 'rb')
    output = PdfFileWriter()
    
    
    folder=os.path.dirname(filepath)
    new_file=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_deleted."+filepath.split("\\")[-1].split(".")[1]
        
    for i in range(infile.getNumPages()):
        if i not in pages_to_delete:
            p = infile.getPage(i)
            output.addPage(p)

    with open(new_file, 'wb') as f:
        output.write(f)
        
        
    print(f"Total {n} pages have been deleted. The New PDf has been stored in same folder \n \n")
    
    
    



def rotate_pdf(filepath,type="normal",degree=0,odd_degree=0,even_degree=0):

    original_file=filepath.split("\\")[-1]
    


    f = open(filepath, 'rb')
    pdf = PdfFileReader(f)
    writer = PdfFileWriter()

    
    
    if type=="normal":
        for i in range(0, pdf.getNumPages()):
            page = pdf.getPage(i)

            page.rotateClockwise(degree)

            writer.addPage(page)

        print(f"All pages of the File {original_file} has been rotated at {degree} degree \n Please wait saving your file...")

    

    elif type=="oddeven":

        for i in range(0, pdf.getNumPages()):
            if (i+1)  % 2 == 0:
                page = pdf.getPage(i)

                page.rotateClockwise(even_degree)
            else:
                page = pdf.getPage(i)

                page.rotateClockwise(odd_degree)

            writer.addPage(page)



        print(f"Odd pages of the File {original_file} has been rotated at {odd_degree} degree and even pages at {even_degree} \n Please wait saving your file...")

    else:
        print("Specify the Correct Type")


    extension = os.path.splitext(filepath)[1]
    filename = os.path.splitext(filepath)[0]
    folder=os.path.dirname(filepath)
    newfile=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_rotated."+filepath.split("\\")[-1].split(".")[1]
    


    output = open(newfile, "wb")

    writer.write(output)

    output.close()
    f.close()

    print(f"The File has been rotated and saved as {newfile}")






def excel_to_pdf(filepath):
    
    
    """
    This function will convert the excel file into the Pdf File
    
    Please ensure that the content in the excel is within the printable area
    
    only one parameter needs to be given i.e the complete file path to the excel file
    
    
    """

    
    original_file=filepath.split("\\")[-1]
    
    excel = client.Dispatch("Excel.Application")

    sheets = excel.Workbooks.Open(filepath)
    work_sheets = sheets.Worksheets[0]
    
    
    folder=os.path.dirname(filepath)
    new_file=folder+"\\"+filepath.split("\\")[-1].split(".")[0]+"_pdf_converted"
    
    work_sheets.ExportAsFixedFormat(0, new_file)
    
    print(f"The Excel {original_file} has been converted to {new_file}")


    
    
def pdf_to_word(filepath):


    with open(filepath, mode='rb') as f:

        reader = PyPDF2.PdfFileReader(f)

        page = reader.getPage(0)

        txtfile=os.path.dirname(filepath)+"/ConvertedWord.docx"
        file = open(txtfile, 'w',encoding='utf-8')
        file.write(str(page.extractText()))
        file.close()
        print("Done!")
        
        
        

        

def ppt_to_pdf(filepath):
    
    
    """
    This function will convert each slide of your PPT into a PDF file .
    
    Only one paramenter needs to be gievn. i.e the Complete Path to the PPT File
    
    The output Pdf file will have the same name as Input file and Will be stored in the same folder.
    
    
    """
    
    print("Your Input file is at:")
    print(filepath)
    


    outputFilePath = generateOutputFilename(filepath);

    print("Your Output file will be at:")

    print(outputFilePath);
    
    print("We are converting Your file, please Wait...!!")


    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

    powerpoint.Visible = 1

    slides = powerpoint.Presentations.Open(filepath)

    slides.SaveAs(outputFilePath, 32)

    slides.Close()

    powerpoint.Quit()

    print("The PPT has been converted to Pdf file Successfuly!")

            

      
def word_to_pdf(filepath):
    
    
    """
    A very simple funtion to covert the word file into pdf file
    
    This will take just one parameter i.e the file pah
    
    """
    
    from docx2pdf import convert
    
    print(f"The file {filepath} has been selected.")
    
    convert(filepath)
    

    print("The conversion has been completed.")
    
