# Robot Framework XLSX Support
## Introduction
 This Custom library will could help you to do the task robotframework with XLSX format . 
 This is still  in Alpha Version. I'll update more function.

## Installation
  Install openpyxl first .
  ```bash
      pip install openpyxl
```
 .and copy robotframework_library_xlsx_alpha.py to your project .
 
   ```bash
      Library           		  lib/robotframework_library_xlsx_alpha.py
```


```robotframework

      *** Settings ***
      Documentation               This is just Test
      ...
      Library           		  lib/robotframework_library_xlsx_alpha.py
      Library                     Selenium2Library
      Library                     Collections
      Library                     OperatingSystem
      Library                     BrowserMobProxyLibrary
      Suite Setup                 Start Browser
      Suite Teardown              Close Browsers
      
      
      *** Variables ***

      
      *** Test Cases ***
      Check something
          TC-02-Generate-User-Payroll-xlsx
			 [Documentation]  Writing Data into excel
			 ${gender_arr}=    Create List
			 ${i_lenght}=       get length      ${usercard_arr}
			 :For    ${i}    in range     0       ${i_lenght}
				 \  ${gender}=       set variable    Male
				 \  Append to List     ${gender_arr}      ${gender}
			 add value by list  tmp_payroll_user.xlsx   2   1   ${gender_arr}
```
