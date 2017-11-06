# SBX3TO2

Tool to export SBX files created by SBuilderX 3.15 in a format that SBuilderX 2.05 can import. It can export back Maps, Lines, Polys, Objects And Excludes. The TYPE parameter of Lines and Polys is written as a null, otherwise SBX205 crashes when reads types used by SBX315. Also the parameter GUID of Lines and Polys is not exported. 

The file SBX3TO2.EXE has been added to this repository and you can run it directly without compiling the project. When you import the file created by SBX3TO2 using SBuilder 2.05 or 2.06, you need to be sure that your decimal separator symbol is a period (.). You can change the decimal separator symbol through the Windows Control Pannel under the Regional Settings.

For more information about SBuilder and SBuilderX please visit www.ptsim.com
