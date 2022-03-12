# Binary-files-inside-EXCEL-VBA

This application can convert any executable file (<kbd>.exe</kbd> or any type of file in fact) to VBA source code that can be included as a module in an EXCEL file. Once inserted into the EXCEL file, the VBA code can be used to completely restore the executable file to disk in the same directory as the EXCEL file. Also, once recomposed on disk, the executable file can be executed automatically. Use the EXCEL file in the "test" directory for testing. Press ALT + F11 to open the VBA window so you can see the code being discussed.

# The EXCEL .xlsm file

Once you press the button on the left, EXCEL will execute the VBA source code from the main module, that was initially produced by the software below (which can be found in the same directory). The VBA source code of the module contains the hexadecimal version of a test binary file. This code execution will use the hexadecimal code stored in an array variable to recreate the original binary file. Once the binary file is made, it is immediately executed and a window that contains the text "I was in the HEX format" can be seen. Note this is a conversion method used also by some malware in the past. However, the method is extremely useful and may help with many hybrid applications of in the case of protocols in between different environments. Thus, some antivirus engines might confuse this method with some malware file.

Press Alt+F11 to see the VBA source code of this EXCEL file.

# Why is this important?


<img src="https://github.com/Gagniuc/Binary-files-inside-EXCEL-VBA/blob/main/img/scr.png?raw=true" alt="">
<img src="https://github.com/Gagniuc/Binary-files-inside-EXCEL-VBA/blob/main/img/excel%20vba%20hex%20to%20bin%20(2).png?raw=true" alt="">
<kbd><img src="https://github.com/Gagniuc/Binary-files-inside-EXCEL-VBA/blob/main/img/excel%20vba%20hex%20to%20bin%20(1).png?raw=true" alt=""></kbd>
