# RTC5Upgrade_Readme

[p.1]
© SCANLAB AG 2009 
RTC5Upgrade  
 
 
 
RTC®5 Upgrade 
The RTC5Upgrade.exe program allows enabling one or more options of a specific RTC®5 
board.  For each board, a separate file RTC5Upgrade_SN_<NN>_<Opt> determines 
which options will be enabled.  <NN> denotes the serial number of the RTC®5 board to be 
upgraded.  <Opt> denotes the option(s) to be enabled. 
How to perform the upgrade: 
Important: The RTC®5 board with serial number <NN> must be installed and must not be 
acquired by another application. It is not necessary that the board is initialized by 
load_program_file. 
1. Copy RTC5Upgrade.exe and the file RTC5Upgrade_SN_<NN>_<Opt> into the 
directory containing RTC5DLL.dll.  
Caution: Do not change the file name nor the content of the file 
RTC5Upgrade_SN_<NN>_<Opt>. Otherwise the upgrade will be discarded. 
2. Start RTC5Upgrade.exe. It displays an explorer window.  
3. Choose the upgrade file wanted.  
The upgrade will be performed followed by a success message.  If discrepancies are 
detected (for example, if the RTC®5 board is not installed, is acquired by another 
application or the EEPROM can't be programmed successfully), a message box with 
an appropriate error message will be displayed.  Then the program terminates without 
upgrading.