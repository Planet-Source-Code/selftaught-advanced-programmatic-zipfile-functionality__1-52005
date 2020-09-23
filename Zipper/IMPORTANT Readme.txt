To run the GrpZipper group project, you will need the following files registered and/or installed on your machine:

AsyncZip.exe    - My wrapper program around Infozip's DLL's
RichDialogs.dll - My messagebox program for custom messagebox and input box used by the test program for getting a password and confirming file overwrite with four options.
WinsubHook2.tlb - Paul Caton's unbelieveably functional TLB.  If you don't have it already, you need it.  Search for Paul Caton on PSC, and wade through the piles of useful code.
Unzip32.dll  
Zip32.dll       - Infozip's Dlls.  Open source, same compression used by Winzip.  See www.Info-Zip.org.


All of these files are available from the zip.  Here are detailed instructions

1. Open the project AsyncZipper\AsyncZipper.vbp, and compile it to a directory of your choosing.  This will automatically register it.

2. Open the project MessageBox\MessageBoxEx.vbp, and go to the Project Menu-> References.  Choose the Browse Button, and Browse to OPC\TLB\WinsubHook.tlb, and select that file.  Once the reference is set, compile RichDialogs.dll to a directory of your choosing.

3. Open the project DecodeDll's and run it.  This will decode the dll files into your system directory, asking first of course!  This is a neat trick I devised to pass the DLL files through PSC's monitor program by encoding it in Base64.  This also serves a good purpose of making the DLL file more compressible.  Feel free to use this project yourself to do the same.

4. open the grpZipper.grp file in the main directory, and you should be off and running!  In case you're not, read below.




If you get VB's error "Cannot find project or Library"

When you compile the AsyncZip.exe, VB will issue it a new set of GUID's, which will probably mean that VB will not be able to find the correct reference in Zipper.vbp.  Assuming that you have compiled to program as in step 1 above, just open up the references dialog and clear the MISSING reference to AsyncZip.exe, press OK, open the references dialog again and set a new reference to the AsyncZip.exe file that you have created.  If your machine requires you to do this, you will probably also have to do it to RichDialogs.dll.

