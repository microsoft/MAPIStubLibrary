# Building the MAPI Stub Library

To build the MAPI Stub Library as MAPIStubLibrary.lib, do the following:

1. Download and unzip the [source](https://github.com/stephenegriffin/MAPIStubLibrary).

2. Download and install the [Outlook 2010 MAPI Header Files](http://www.microsoft.com/downloads/en/details.aspx?FamilyID=f8d01fc8-f7b5-4228-baa3-817488a66db1&displaylang=en).

3. Open the project for the MAPI Stub Library matching the version of Visual Studio you have. We provide project files for Visual Studio 2008 and Visual Studio 2010.
   - **Visual Studio 2008**: The project file is located in the vs2008 directory. After opening it, go to Tools>Options>Projects and Solutions>>VC++ Directories, and add the directory for the Outlook 2010 MAPI headers prior to the Visual Studio include directories.
   - **Visual Studio 2010**: The project file is located in the vs2010 directory. After opening it, right click on the MapiStubLibrary project and select Properties. Switch to the Configuration Properties>VC++ Directories node and add the directory for the Outlook 2010 MAPI headers prior to the Visual Studio include directories.

4. From the Release Configuration drop down, select Release, unless you're building a library for debugging purposes.

5. From the Solution Platforms drop down, select Win32 or x64, depending on the flavor you need for your application.

6. Build the project.

You now have a file, MAPIStubLibrary.lib, which you can link in to your project.