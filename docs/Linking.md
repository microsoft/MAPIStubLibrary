# Linking To MAPIStubLibrary.lib

You can use the MAPI Stub Library as a drop in replacement for mapi32.lib. Once you've built MAPIStubLibrary.lib, you can remove mapi32.lib from your linker settings and replace it with MAPIStubLibrary.lib. No further modifications to your code should be necessary.

ExampleMapiConsoleApp demonstrates this technique. Note that ExampleMapiConsoleApp.cpp includes the regular MAPI headers and has no code dedicated to loading MAPI dlls. When the first MAPI call is made, in this case, MAPIInitialize, the stub library takes care of looking up the location of Outlook's implementation of MAPI so it can load it and dispatch the call. Subsequent MAPI calls continue to use the same dll, lining up entry points with GetProcAddress.

For more information, see [Building the MAPI Stub Library](Building.md).