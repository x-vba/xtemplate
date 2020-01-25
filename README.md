# XTemplate

## Description

[XTemplate](http://x-vba.com/xtemplate) is a tool for Microsoft Word, PowerPoint, and Outlook that is used to
create templates that can easily pull information from Excel Workbooks so that
you don't have to fetch this data manually. It provides a very simple template
syntax, and makes working with recurring, standardized Documents, Presentations,
and Emails much easier.

## Usage

To use XTemplate, simply put templates throughout your Document, Presentations,
and Emails, and then run the XTemplate macro. If the syntax is correct, and the
Workbooks you want to fetch data from exist within the correct folder, the templates
will be replaced with the value in the respective Workbook. The template syntax looks
like this:

{{ C:\Files\\\[MyWorkbook.xslx]MySheet!A1 }}

In this case, when running the XTemplate Macro, this template will be replaced with
the value in Range A1 within a Sheet named MySheet, within a Workbook named
MyWorkbook.xlsx, which is founded within the folder C:\Files\

## Download and Installation

For more information about the template syntax, downloads, and installation,
please see the official documentation.

## License

The MIT License (MIT)

Copyright © 2020 Anthony Mancini

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
