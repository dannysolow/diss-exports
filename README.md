# diss-exports
VBA macro to fix DISS export formatting issues

**Introduction**
The combination of merged cells and mis-aligned data in DISS makes the export very difficult to use. Any new export requires considerable cleanup. To help with that, this is a simple Excel Macro to reformat the contents as a normal readable file. 

This repo includes both a .xlsm file with the macro already enabled as well as the underlying Visual Basic (VBA) code which you can use. For those who have restrictive security policies or are wisely are untrusting of macro scripts from the internet, you can simply plug the VBA directly into your own Excel.

![](diss-improve-preview.gif)

**Disclaimer**
This macro is experimental. Any data you enter into this excel-enabled workbook should be backed up. While potential impact is minimal, use at your own risk.  

**Instructions**
1. Either download the .xlsm file or copy and paste the VBA into your own macro-enabled excel file.
2. Copy and Paste your DISS export In Cell A:1. It is important you overwite all the content here including the instructions
3. Enable macros: Go to the File tab > Options. On the left-side pane, select Trust Center, and then click Trust Center Settings… . In the Trust Center dialog box, click Macro Settings on the left, select Enable all macros and click OK.
4. Go to the "View" tab in the toolbar.
5. Click on "Macros" and select "View Macros."
6. Run the DISS_Reformat macro
