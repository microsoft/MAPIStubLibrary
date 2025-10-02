# Directly Including the MAPI Stub Library

One way to incorporate the MAPI Stub Library is to copy the source files, MapiStubLibrary.cpp and StubUtils.cpp directly into your project and remove any linkage to mapi32.lib as well as any code which explicitly links to MAPI. [MFCMAPI](http://mfcmapi.codeplex.com/), which formerly implemented [Explicit Linking](http://msdn.microsoft.com/en-us/library/cc963763.aspx) to link to MAPI, uses this technique. If you're not already including msi.lib, you will need to either include it or manually load the MSI functions.

Since the MAPI Stub Library implements all of the logic for loading MAPI for you, in MFCMAPI we were able to eliminate a large amount of code from ImportProcs.cpp which had previously handled locating MAPI and finding entry points to MAPI functions. A couple features in MFCMAPI changed during this integration, such as:

- **Forcing MAPI to load from a user input path**: MFCMAPI can still do this, by loading the DLL as directed by the user, then using the SetMAPIHandle function to inform the stub which module to use.
- **Manually unloading MAPI**: The stub library implements a routine, UnLoadPrivateMAPI, which MFCMAPI now uses to unload MAPI.

Aside from including the source files as part of your build and adding msi.lib, there are no changes required to your code which makes MAPI calls.