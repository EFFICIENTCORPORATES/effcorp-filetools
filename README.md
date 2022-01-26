# Efficient Corporates File Tools Utilities in Python
First of all, Welcome to Efficient Corporates..! 
Before you start reading the Documentation, let us please give you a small Disclaimer


*At, Efficient Corporates (or Eff Corp as we call it), we believe that Coding is for everyone and 
everyone has the right to make their life simpler through use of automation.
Hence, we endeavor to make it simpler for non coders (like the people without Computer Science as their Major
or even Commerce / Arts  background people)to make coding a daily habit . Since,
we believe on making things less technical and more practial, hence, it is possible that 
we may not follow the best coding practices or documentation practice to ensure that 
users get to understand easily , So.. PLEASE EXCUSE US  for that*



### Lets know a bit about Efficient Corporates

Efficient Corporates is an Open Knowledge Sharing Platform which works towards 
encouraging the HR, Sales , Opertions , Finance , Tax , Accounts or any other Professionals to 
adopt technology in their day to day office Working and automate their most boring repetitive
tasks at Office (so that they can have a meaningful 9 to 5 job or WFH)

We have come up with our Second Module on FileTools Automation and we plan to come up with many more such 
modules.

### (Our First Module was effcorp-gst. [This is primarily for the practiciing Tax Professionals in India doing GST related work]
If you have not checked that out yet, we highly recommend you to to check that out, if you are someonw who work in GST)


[Link to the Module effcorp-gst](https://pypi.org/project/effcorp-gst/):

You can also watch here the [Complete Video Tutorial on effcorp-gst](https://www.youtube.com/playlist?list=PLaso8-OZjhbx9y5QuaNhVs-95_r6rqzP4):


### Now, lets get back to this Module "effcorp-filetools"

This Module of filetools (is just like any any other python module like Pandas , Numpy, matplotlib) and has
some functions (pre defined actions) which help us to automate the files related tasks.

*Again, if you are very new to Python, and have no Idea what's going on here, we strongly recommend you to 
first go through our very basic Python Tuitorial Especially designed for Finance Professionals and students*

In Our Youtube Channel "Efficient Corporates" [Python Tutorial for Absolute beginners by Effcient Corporates](https://www.youtube.com/watch?v=E509BVUxrZg&list=PLaso8-OZjhbyTgqcLSxbusK2RpPu_c3lC):


Now, what task does this module actually help us perform?

Below are the major utilities present in this module:

### Utilities Present in the Current Latest Version of Eff Corp Filetools

1. Search for Files in your Computer/ Laptop   
2. Excel Related Tools
3. Pdf Related Tools  
4. Converting from one file format to another file format


So, these are the broad heading of the  functions this module performs.

Let's read about each of these in details



## Installation

As we mentioned that effcorp-filetools is like any other Python Modules like pandas, matplotlib, numpy,
so, even this can be installed using the simple pip command as below. (Documentation available at [PyPI](https://pypi.org/project/effcorp-filetools/):)

```pip install effcorp-filetools```

The Efficient Corporates Filetools will better run on Python 3.6 and above. 
It is so because , we have used string formatting f literals which work only on python 3.6 and above


## How to use?

Now, lets see how do we use Pandas . We simply pip install it and then import the module saying


```import pandas as pd```

But, there are some modules like Tkinter, where we write as 

```from tkinter import *```

So, in this case , though we have pip install effcorp-filetools , but the entry point to this module is
set to the keyword "gst".

This means you can enter inside this module using the Keyword "filetools" only.

Quite weird, but that is the way the module has been set up, So we will need to do the below to get inside the module

```import filetools```

##### OR

```from filetools import filetools_utilities```


***Below codes to import will not work***

```import effcorp-filetools``` >>> Won't work


```from effcorp-filetools import filetools``` >>> Won't work




### Utilities Under this File Tools (Lets talk about thse One-By-One)

### 1. Search for Files in your Computer/ Laptop 

Many times, we come across a situation where we want to search for certain file.

Or May be search for files with particular extension.

So, won't it be wonderful , if you would get the names and the complete location of that files in an excel format

To make this easier, we have come up with 2 module depending upon your requirement.


#### Requirement 1:

##### 1.1 Searching particular extension in entire Laptop/ Computer System

Name of Function : **search_drives()**

```buildoutcfg
from filetools import filetools_utilities

filetools_utilities.search_drives(".pdf")

```
Now, if you think that just tooo Long, you can shorten it like below:

```
from filetools import filetools_utilities as ftu

ftu.search_drives(".pdf")

```


Just with these 2 lines of code, the program will search your entire system and find out the
.pdf files in your system (i.e it will look inside your C drive, D Drive, E Drive, F drive and whatever drives your have)

#### Just a small caution that, since thse will look for the extension in your entire system So these might take a bit longer time to run. So you will have to have that patience.

This program takes around 2 min - 10 min depending uoon the number of files you have in your system

##### Output:
You will get an Excel file with beow 3 columns: (Excel will be stored in that same folder from where the code is run) 

a. Name of the File of that Extension (i.e all pdf files in this case)

b. Complete Location of that file

c. Size of that File (in Kb)


Now, what if you want to search only the D Drive , or may be only inside a particular folder only , which is inside C Drive.

So, you will need the Requirement 2:

#### Requirement 2:

##### 1.2 Searching particular extension file in particular Folder and inside its Subfolder


Name of Function : **search_folder()**


```

from filetools import filetools_utilities as ftu


ftu.search_folder(r"C:\Users\Dell\Documents",".pdf")

```

***


#### Requirement 3:

##### 1.3 Searching particular extension file in particular Folder and inside its Subfolder BUT also fulfills certain name conditions


Name of Function : **search_files_condition()**

```

from filetools import filetools_utilities as ftu


ftu.search_files_condition(r"C:\Users\Dell\Documents",".pdf","case_sensitive"=False)

```

In this we will specify the path In which we want to look for the particular extension file
Also, on running this, we will get a input box, where we ca specify the pettern in regex format.

Like, if we want to find all files starting with "E", then we can simply give
E.  in the input box. The . will mean that any number of character after E

For a complete regex documentation, you can also consider reading the Python Documentation

Link : https://docs.python.org/3/howto/regex.html

***



### 2. Excel Related Tools

##### 2.1 Splitting Excel
In Finance, we have been working with excel files since a very long time.
So, many time we come across a situation in which we need to divide the Excel Files into different files or into multiple shets.
Usually , this splitting is done based on certain columns (like Product wise, Department Wise) which is present inside that Excel File.

##### 2.1.1 Splitting into Different Excel Sheets in single excel file

Now, in this code, you will just need to specify the Excel Files and the columns name, 
which you want to Split and also the Column Name basic whose value you want to Split the data

```
from filetools import filetools_utilities as ftu

ftu.split_excel(filepath=r"F:\New Folder \MODULES\Important\Financial Sample.xlsx",
                cols="Month Name",
                sheet="Main",
                mode="sheets")
                
```

**NOTE**
In the above code, the function will read the file in that filepath variable.

It will read the sheet named "Main" and read the table in that Excel file.  Then, it will  select the column whose heading is "Month Name"

The Mode is selected as "sheets", so for each value in that column Named "Month name", a sheet will be created

So, here 12 sheets for each month will be created in a single file.


##### 2.1.2 Splitting into Different Excel files


So, if you want to create different Excel files, instead of sheets, so use the mode as "file" 

```
from filetools import filetools_utilities as ftu

ftu.split_excel(filepath=r"F:\New Folder \MODULES\Important\Financial Sample.xlsx",
                cols="Month Name",
                sheet=1,
                mode="file")
```

The code is same as above, just the mode is different.

Also, note that the sheet is mentioned as 1.

This means that the second sheet will be read (Python starts counting from Zero , so 1 means the second sheet)


### Next Utility in Excel

This utility is around the combining of the Excel files

##### 2.2 Combining Excel 

If we have same data divided into multiple excel sheets or in multiple excel files,
then we can use this program to merge the data.

##### 2.2.1 Combining the different Excel files

Suppose we have a same sales ledger extracted from SAP or from Tally, for each month in separate excel files and we 
want to consolidate it into a single excel file, then we can make use of this command.


```
from filetools import filetools_utilities as ftu


ftu.combine_excel(folder = r"F:\New Folder \MODULES\Important\ Excel Files",
                mode="file",
                sheet=0)

```
The above code will read all the excel file in that folder named "Excel Files"

Since the mode is mentioned as "file" it means that we have data in differet excel files.

Also, sheet=0 , means that  program will read the first sheet of each excel file.

If we want program to read 4th Sheet of each excel file, then we need to mention sheet=3

Alternatively, you can also give the name of the sheet like "Data" or "Master".

In this case, the name of the sheet must be same in all the excel files. 

After reading as per the parameters mentioned, the program will give a single sheet as Output.

**NOTE** 
Here, If you want to merge Excel files , you DO NOT need to provide the complete filepath , You just need to give the folder path.
This program will automatically read all excel files which are in that folder.
Please note that all the files must have the exact same headings.

If the heading is different e.g "Month" & "Month Name", then this program will create keep two columns "Month " and "Month Name"

So, all the headings should be exactly same. 

Order / Sequence of the columns is not Important.



##### 2.2.2 Combining the different Excel Sheets in same Excel file



Next, lets Suppose we have a same sales ledger extracted from SAP or from Tally, for each month in separate sheets but within a same excel file,

Now, if we want to consolidate it into a single excel file, then we can make use of this command.


```
from filetools import filetools_utilities as ftu


ftu.combine_excel(folder = r"F:\New Folder \MODULES\Important\ Excel Files\ The Excel.xlsx",
                mode="sheets")

```


***

**NOTE**

So, here you need to note the below points:

1. In folder = r''' , we have given the complete Excel file path and not just the folder path

2. Mode need to be given as "sheets"

3. The parameter sheet=0 , need not be given as this program will read all the sheets in that 1 excel file.




### 3. Pdf Related Tools

Here we will be looking at multiple tools, which let us do the various tasks of office 
with just a single line of code


#### 3.1 Combining Pdf

##### Variant 1: Combining pdf one after another

It is a very easy function , and can be done by using most of the free sites online.

But, if you want protect the confidentiality of your data, then better do it on your own. In your local system, without uploading it online

So,below is a program that you can use to combine the pdf files in seconds without need to upload it in internet

```

from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files"

ftu.combine_pdf(filepath=path)

```


##### Variant 2: Combining pdf odd even pages

In this case, if we have multiple pdf files, some have od number of pages and some have even number of pages.

Bw, if all such files are combined as it is (as we discussed in Variant 1), then, if we do the both side 
print of the combined pdf, then first page of the pdf after the odd pages pdf will appear at the backside
So, to prevent this, we have this function 

This will add a blank page if the total number of pages in the pdf file is odd numbered.

In this was the both side print out of the consolidated pdf file can be easily taken.
So,below is a program that you can use to combine the pdf files in seconds without need to upload it in internet

```

from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files"

ftu.combine_pdf_oddeven(filepath=path)

```

***
#### 3.2 Splitting Pdf

This module is used for the splitting of the excel file.
Splitting can be done in many ways

Code for Splitting the pdf file will be as below:


There are Multiple ways of Spliting the excel file

**3.2.1 type="page_wise"** 

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.split_pdf(filepath=path,type="page_wise",set=1)

```

In case of a pdf with 16 pages is splitte using type "pages wise", 
it will give 16 pdf with 1 pages each.

First pdf will have page 1
Second pdf will have page 2
Third pdf will have page 3 ..and so on...


**3.2.2 type="cummulative"**

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.split_pdf(filepath=path,type="cummulative",set=1)

```
In this case, 

first pdf will have page 1 page
Second pdf will have Page 1 & 2
Third Pdf will have Page 1 & 2 & 3
.
.
.
15th Pdf will have Page 1-15
Last pdf will have All 16 pages

So, in this way the pages keeps on accumlating. In this manner, the 16 paged pdf  will be splitted into 16 files.


**3.2.3 type="oddeven"**

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.split_pdf(filepath=path,type="oddeven",set=1)
```

The odd pages will be splitted in one file and the even pages will be splitted in another file


**3.2.4 type="ranges"**

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.split_pdf(filepath=path,type="ranges",set=5)

```
Under this, we need to specify, at what range/ interval do you want to break.

For Example: If a pdf has 16 pages, and the type is selcted as ranges and the Set is given as 5.

So, it basically tell the system to make pdf files with 5 pages each.

So, since the pdf has 16 pages, the program will make total 4 files

First 3 will have 5 pages each and the fourth one will have 1 page i.e the remaining page




**3.2.5 type="split_equal"**

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.split_pdf(filepath=path,type="split_equal",set=4)
```

In this cases, the program will return 4 pdf files with equal page numbers

If a pdf have 16 pages , this will return 4 pdf with 4 pages each



##### If you want any other kind of split, do provide your feedback.
##### We will try to include that in our future versions

***

#### 3.3 Sorting Pdf

This function will reverse the order of the pdf file

Last page will be the first page ,  second last page will be second page.... and first page will be the last page

```
from filetools import filetools_utilities as ftu

filepath=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.sort_pdf(filepath)
```

Currently,we have only one type of sorting i.e reverse, in future versions , we may come up with more
sorting oprtions (Feel free to give your feedbacks on efficientcorporates.info@gmail.com)

***

##### 3.4 Encrypting Pdf

**3.4.1** Encrypting the pdf files with the default Password

In the below code, we have given only provided the path of the file, so the default password ie. Password@1 will be the password.

The pdf file will get encrypted with this password and saved as a new file.

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.encrypt_pdf(filepath=path, type="function")
```

**3.4.2** Encrypting the pdf files with the different Password

In the below code, we have specified the password within the function, hence the pdf file "new_File.pdf" will be enrypted with the password "Effcorp@11"

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.encrypt_pdf(filepath=path, password="Effcorp@11", type="function")
```


**3.4.3** Encrypting the pdf files with user Defined password

NOw, if we want to ask the user , the password he wants to set for the pdf file,
in that case, we need not give the password and can set the type as "user"

On running this code, the pdf file new_File.pdf, will be selected

Then user will be asked to enter the password he wants to set for the said pdf file.

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.encrypt_pdf(filepath=path, type="user")
```
***

##### 3.5 Decrypting Pdf

**3.5.1** Decrypting the pdf files with the Password given inside function

**Let us first make it clear that, here we are not cracking the password.

We are simply decrypting the pdf file with password provided.

Pasword must be known by you and this function will simply remove the password and save the file as a normal pdf file with no password**

So, writing the code will be as given below:

In the below case, we have  provided the path of the pdf file which is to be decrypted.

Further, we have also provided the password for decrypting
```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File_encrypted.pdf"

ftu.decrypt_pdf(filepath=path, password="Finance@11", type="function")
```

This will save a new file in same folder without any password

***

##### 3.6 Numbering the Pdf File

This will add page numbers to your pdf file.

Simply the path of the pdf files needs to be mentioned.


```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"

ftu.addnumber_pdf(filepath=path)
```

***

##### 3.7 Delete Selected Pdf Pages

The code will be as below:

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"


ftu.delete_pdfpage(path)

```
On running this code, it will ask you for the below:

**The pdf file new_File.pdf has been selected**

**Enter no.of pages to delete :**

So, here you need to provide how many number of Pages to delete. ike 5 or 7 or 10

After this, it will ask you, which pages to delete...

**Enter page numbers(with spaces no commas) :**

So, if you want to delete Page num 4, page 6 & page 9

You need to give the input as 4 6 9 

(Do not give any comma or hiphen, simply type the pages number with spaces)

***

#### 3.8 Rotating Pdf

**3.8.1 Rotating all pages in certin direction**

You need to provide the path of the pdf file, and simply provide the degree at which rotation is to be done.

Rotation degrees to be given in multiple of 90

Rotation will take place "Clockwise"

```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"


ftu.rotate_pdfpage(filepath=path,type="normal",degree=180)

```


**3.8.2 Rotating odd pages in certin direction & even pges in other direction**


```
from filetools import filetools_utilities as ftu

path=r"F:\New Folder \MODULES\Important\ Pdf Files\new_File.pdf"


ftu.rotate_pdfpage(filepath=path,type="oddeven",odd_degree=-90,even_degree=180)

```

***




### 4. Conversion from One File Format to Another



#####  Excel to PDF Conversion

So, if there are multiple excel files that you want to convert pdf to , you 
can very well make use of this function and run your files in a loop to convert all your 
excel files into pdf


```
from filetools import filetools_utilities as ftu


Complete_file_path=r"F:\New Folder \MODULES\Important\ Excel Files\Excel_file.xlsx"


ftu.excel_to_pdf(Complete_file_path)

```
Note: You will have to ensure that the excel file has the contents in
a printable manner.

Print area must be set in the excel file.


#####  Word to pdf  Conversion

Using this function , you can easily convert any number of word files into a pdf files


```
from filetools import filetools_utilities as ftu


Complete_file_path=r"F:\New Folder \MODULES\Important\ Word Files\My_file.docx"


ftu.word_to_pdf(Complete_file_path)

```


#####  Pdf to Word Conversion

Many times you will need to edit a pdf file. NOw, it may not be possible
to correctly convert a scanned document , but if the pdf file 
is clear enough, we can have the pdf converted to word quite fairly.


```
from filetools import filetools_utilities as ftu


Complete_file_path=r"F:\New Folder \MODULES\Important\ Pdf Files\My_file.pdf"


ftu.pdf_to_word(Complete_file_path)

```



#####  PPT to PDF Conversion

If you want to convert a power point file into a Pdf file, instantly, you can
use this fucntion to do so..!

If you have multiple files, then you can create a loop to un through
all the ppt files and then create pdf file for each one of them.


```
from filetools import filetools_utilities as ftu


Complete_file_path=r"F:\New Folder \MODULES\Important\ Presentation Files\My_file.docx"


ftu.ppt_to_pdf(Complete_file_path)

```




## License
Since, you have made the effort of reading the documentation till here, let me also explain in simple terms
what this license is all about.

Basically, this code is under a License GNU GPL, which basically means that you are free to use this code in your 
personal use or even use in your office.

And , interestingly, you can even give this code to someone else and also use this cde as a dependency in your own project

Preety much You can do everyting...But....

What you cannot do is to sell this code, or any of your project which uses this code with a commercial interest.

The Bottom Line is "Anything which you got for free, should be available freely..!!"
© 021 Efficient Corporates 

This repository is licensed under the OSI Approved :: GNU General Public License v3 or later (GPLv3+). See LICENSE file for details.


##For any issues / suggestions / complaints/ feedbacks / error faced / or even if you simply want to connect, we have our all ears for you...!! 

##JOIN our Community "EFFICIENT CORPORATES" NOW


[Discord Server](https://discord.gg/MB7PWfpau3):

[Youtube Channel](https://www.youtube.com/c/EFFICIENTCORPORATES):

[LinkedIn Company Page](https://www.linkedin.com/company/efficient-corporates/):

[LinkedIn Discussion Group](https://www.linkedin.com/groups/13967995/):

[Quora Space](https://efficientcorporates.quora.com/):

[Facebook Page](https://www.facebook.com/efficientcorporates):

[Twitter Handle](https://twitter.com/EfficientCorp):




