# ExcelFunctionRegex

## Description
***ExcelFunctionRegex*** - an extension for Microsoft Excel written using the *Excel-DNA* library, which adds to Excel the ability to check, search and replace by a regular expression.

## Installation
Extension files include XLL files located in the directory [ExcelFunctionRegex/bin/Release/](ExcelFunctionRegex/bin/Release/). It is clear that 64, in the file name, means 64-bit Excel, but *packed* means that any external files are packed into this file. In this case, no additional files were used when creating the extension, so there is no difference which of the files to use packed or not, only the bit depth matters. 
They should be installed as follows:
- In the Excel parameters, select the tab "*Надстройки*"; 
- At the bottom of the Excel parameters window, with the selected field value "*Надстройки Excel*", click "*Перейти...*". A small window will appear with a list of all possible add-ons in Excel.;
- You need to select any of the add-ons, and click "*Обзор...*", at the same time, it is not necessary to connect this add-on at all;
- A dialog box opens in which you can see the path where the add-ins should be located. You need to remember this path, but it's better to copy it to the buffer;
- It is necessary to close all Excel dialog boxes, and Excel itself should also be closed;
- You need to open any folder and use the remembered path to go to the directory where Excel extensions are stored, although when you open this directory, it will most likely still be empty;
- В этот директорий надо скопировать один из файлов расширение из каталога [ExcelFunctionRegex/bin/Release/](ExcelFunctionRegex/bin/Release/), in accordance with the bit depth of the installed office;
- After the extension is in the folder, you need to open it again "*Параметры Excel*" and in the tab"*Надстройки*" with the button "*Перейти...*" open a window with a list of add-ons. This time you will already be able to see the added extension in the list of add-ons *ExcelFunctionRegex Add-In*, which must be activated by checking the box next to it.

## Using
In the list of function categories, find and select the *ExcelFunctionRegex Add-In* section, after that you will be able to select one of the functions:
- **ДелТекстПоРегВыр** - function text separation by regular expression;
- **ПоискПоРегВыр** - regular expression search function;
- **ПроверкаПоРегВыр** - regular expression validation function;
- **СцепТекстПоРазд** - combining text through a separator character.

You can learn more about the *ExcelFunctionRegex* extension, its installation and use, as well as about creating your extensions and additional functions using the *Excel-DNA* library on my website "[Парадокс-Портал/Расширения для Microsoft Excel на C#. Дополнительные функции регулярных выражений.](http://www.paradox-portal.ru/blog/article/9-rasshireniya_dlya_microsoft_excel_na_c_sharp_funkcii_regulyarnyie_vyirazheniya)"
