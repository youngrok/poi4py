# poi4py - Apache POI binding for python.

Python has various excel libraries like openpyxl, xlrd, pyexcelerate, etc. They are easy, pythonic, good enough for some tasks, but all of them have their own weakness. Some can only write excel, some can only read excel, some can handle only old excel format or only new excel format. Especially none of these can handle password encrypted excel file. However, Apache POI can handle all formats of excel, read and write, handle password encryption, and more mature. So I wrapped Apache POI for python.

It supports only python 3.

## Installation

	pip install poi4py

## Usage

```
import poi4py
poi4py.open_workbook('myexcel.xls', 'mypassword')
```
