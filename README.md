# Students Shuffling Tool
SOFTWARE MANUAL

## CONTENTS

- About the SST:
- System requirements
- Usage
   - launching the tool
   - read custom columns from excel file
   - change generated column names in the excel file
   - change output file name
   - Change the shuffling type:
- Table of commands


## About the SST:

### STUDENTS SHUFFLING TOOL

A tool made to read an excel file and generate an excel file with multiple sheets based on the user
input where the user can choose the excel file to read from, number of sections, number of stacks that he
wants the shuffling for.

The tool reads and writes excel files with extension of "xlsx" and does not support the old "xls" version.

### System requirements

Please ensure that your computer meets or exceeds the following system requirements before installing
the SST application.
|Name|description|
|---|---|
|Available disk space |1 MB|
|Operating system| Windows, Linux, Mac|
|Environment Variables|<br>**Python v3.5+**. <br>**Pip list**:- <br> et-xmlfile==1.1.<br>  numpy==1.19.<br> openpyxl==3.0.<br> pandas==1.1.<br> python-dateutil==2.8.<br> pytz==2021.<br> six==1.16.<br> xlrd==2.0.<br> XlsxWriter==3.0.|

### Usage

Please make sure you follow these steps to ensure that the tool runs perfectly on your system:

1. Make sure you have the environmental variables mentioned in the requirements section above
    installed either on your python local environment or the global one.

### launching the tool

2. After you have the environment ready you can launch the tool in help mode to see the available
    commands supported by the tool

```
$ python index.py -h
```
![01](/pictures/01.PNG)
```
After you run the command $ python index.py –h you will see the description of the too in addition to
optional arguments as a list contains all available commands you can perform with the tool
```
3. Inside the same folder I attached 2 samples of excel files, where you can get a clue about the type of
    formatting you should adapt your excel file on.


|Sample1| Sample2|
|-------|--------|
|![02](/pictures/02.PNG) | ![03](/pictures/03.PNG)| 


4. Now what’s left is to use one of these samples or your own excel file to get shuffled output of the file,
    to do that you just need to use the **–** - file or **–** f command to enter a file name in the same directory of
    the app in order to be able to read it

```
$ python index.py --file sample
```
5. If you hit enter now, you will get a file called “output.xlsx” which is the default name and shuffling
    type if the user didn’t specify the type.
    
![04](/pictures/04.PNG)

![05](/pictures/05.PNG)


6. If we opened the generated file “output.xlsx”: 

![06](/pictures/06.PNG)

7. As you can see in the image in the 6th step we have multiple sheets and inside each sheet there are
    the students we shuffled, the shuffling is based on number of sections and number of stacks, where
    each sheet represents a section per stack example:
       a. Web Fundamentals, section1, Web Fundamentals, section2, Python, section1...
8. In the command in step 5, we shuffled the students in “sample2.xlsx” using shuffling type1, which is
    set by default if we didn’t specify the type of shuffling we want from the tool, as for now there are 2
    shuffling types supported by the tool, (shuffling type 1, shuffling type 2).

### read custom columns from excel file

If you want to add your own excel file and read specific columns you can us this command

“-rclsl, --rcolslist” which stands for read columns list, for example Name, Academy ID

Note: make sure to put a space between the words so argparser can understand that this is a list,

Also if you have a word of two parts like **“Academy ID”** make sure to wrap them in double quotation.

```
$ python index.py -f sample2 --rcolslist Name “Academy ID”
```
![08](/pictures/08.PNG)

### change generated column names in the excel file

If you want to add your own excel file and rename the generated columns you can us this command

“-wclsl, --wcolsl” which stands for write columns list, for example ID, Name

$ python index.py -f sample2 –-wcolsl ID Name

![10](/pictures/10.PNG)

By default the column names are indexed like 0, 1 for the shuffling type 1,

![09](/pictures/09.PNG)

And for the shuffling type 2, column names are based on the section number like section1, section2.....

### change output file name

If you want to add your own excel file and rename the generated columns you can us this command

- gexn, --generatedexcelname, which stands for generated excel name.

```
$ python index.py -gexn students_shuffled
```
![11](/pictures/11.PNG)

### Change the shuffling type:

To change the type of shuffling you can switch to it by using the flag “--shuffletype “:

```
$ python index.py -f sample2 --shuffletype 2
```
![07](/pictures/07.PNG)

As you can see the shuffling now is based on each stack and each column represents a section inside
that stack.

For now the too supports only 2 shuffling options:

1- **Stack, Section** per sheet

2- **Stack** per sheet


## Table of commands

I’ll list the table of commands/operations that the tool can perform with the description of each command
and when to use, also note that this list can be accessed from the command line using the first command

```
$ python index.py –h
```
|flag | description|
|---|---|
|**-f, --file**| Enter an existing file name, by default^ set to^"sample1", Reads string.|
|**-p, --path**| Enter a path file system, Ex."c:\users\m2\osaid\sample.xlsx", Reads file system path.|
|**-sct, --sections**| Enter the numbers of sections to shuffle, by default set to "3", Reads integer.|
|**-st, --stacks**| Enter a list of stacks to shuffle Ex. "Web Fundamentals" "Python" "Java" ... , by default set to ["Web Fundamentals", "Python", "Java", "MERN"], Reads string.|
|**-gexn, --generatedexcelname**| Enter the name of the generated excel file Ex."test1", by default set to "output", Reads string.|
|**-wclsl, --wcolsl**| Enter generated column names, by default [] works only with shuffling type "1", Reads multiple arguments of any type.|
|**-sft, --shuffletype**| Enter a number that represents a supported shuffling type as for now there is two shuffling types "1" and "2", by default set to"1", Reads integer.|
|**-wst, --withstyle**| Enter bool for so basic styling, currently this feature is not ready to release :3, Reads boolean (true, false, t , f, 0, 1)|
|**-rcln, --rcolname**| Enter column name to read from, by default "Name" Reads string.|
|**-rclsl, --rcolslist**|  Enter a list of columns to read from 2 columns max for now :((((, by default empty [], Reads multiple arguments of any type.|

