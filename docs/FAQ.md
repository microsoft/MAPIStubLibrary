# FAQ - Frequently Asked Questions

## When to use the MAPI Stub

- [Why would I use this library?](#why-would-i-use-this-library)
- [Can't I just use the 64 bit mapi32.lib that ships in the Microsoft SDK?](#cant-i-just-use-the-64-bit-mapi32lib-that-ships-in-the-microsoft-sdk)
- [I explicitly link to MAPI. Do I need this library?](#i-explicitly-link-to-mapi-do-i-need-this-library)
- [Doesn't this library just load the MAPI stub from the system directory?](#doesnt-this-library-just-load-the-mapi-stub-from-the-system-directory)
- [My application is only 32 bit and uses mapi32.lib. Do I have to switch to this library?](#my-application-is-only-32-bit-and-uses-mapi32lib-do-i-have-to-switch-to-this-library)
- [Do you have a list of functions linked from this library that aren't available in mapi32.lib?](#do-you-have-a-list-of-functions-linked-from-this-library-that-arent-available-in-mapi32lib)

## Linking and building

- [Can I just include the source of this library in my project?](#can-i-just-include-the-source-of-this-library-in-my-project)
- [Where are the instructions for building this library?](#where-are-the-instructions-for-building-this-library)
- [What about instructions for linking to it?](#what-about-instructions-for-linking-to-it)
- [Why do I have to build this? Why can't I just download a pre-built .lib?](#why-do-i-have-to-build-this-why-cant-i-just-download-a-pre-built-lib)
- [I use a compiler other than Visual Studio, or a different version of Visual Studio than one of the provided projects. What can I do?](#i-use-a-compiler-other-than-visual-studio-or-a-different-version-of-visual-studio-than-one-of-the-provided-projects-what-can-i-do)

## Bugs

- [I don't like the way the library does X. Can I change it?](#i-dont-like-the-way-the-library-does-x-can-i-change-it)
- [The MAPI Stub Library doesn't export the function X. Can I add it?](#the-mapi-stub-library-doesnt-export-the-function-x-can-i-add-it)
- [The MAPI Stub Library exports the function X incorrectly. Can I fix it?](#the-mapi-stub-library-exports-the-function-x-incorrectly-can-i-fix-it)

**Didn't find what you're looking for here? Start a thread on the [discussion board](https://github.com/microsoft/MAPIStubLibrary/discussions).**

## Why would I use this library?

For years, MAPI was implemented as a 32 bit API. We provided a library, mapi32.lib, which programs could link to in order to get access to MAPI. Then along came Outlook 2010 and 64 bit MAPI. Since mapi32.lib was a 32 bit library, 64 bit programs needed something else in order to use MAPI. This project fills that gap.

## Can't I just use the 64 bit mapi32.lib that ships in the Microsoft SDK?

No - this is a bad implementation which doesn't locate MAPI properly and doesn't link to some MAPI functions correctly. It should be avoided.

## I explicitly link to MAPI. Do I need this library?

No, but you could look at eliminating your explicit linking code by linking to this library.

## Doesn't this library just load the MAPI stub from the system directory?

No, though that is a fallback mechanism. This library implements the technique discussed in [How to: Choose a Specific Version of MAPI to Load](http://msdn.microsoft.com/en-us/library/dd181963) to find Outlook's implementation of MAPI first. Only if this fails does it try the stub from the system directory.

## My application is only 32 bit and uses mapi32.lib. Do I have to switch to this library?

No, but doing so will give you better support for loading Outlook's MAPI in a wider variety of locales. It will also eliminate the need to write LoadLibrary/GetProcAddress code to handle newer exports which are included in this library but are not included in mapi32.lib.

## Do you have a list of functions linked from this library that aren't available in mapi32.lib?

Here's a partial list:

- [GetDefCachedMode](http://msdn.microsoft.com/en-us/library/ff960261)
- [HrGetGALFromEmsmdbUID](http://msdn.microsoft.com/en-us/library/ff522804)
- [HrOpenOfflineObj](http://msdn.microsoft.com/en-us/library/ff960538)
- [MAPICrashRecovery](http://msdn.microsoft.com/en-us/library/ff960295)
- [OpenStreamOnFileW](http://msdn.microsoft.com/en-us/library/gg318092)
- [WrapCompressedRTFStreamEx](http://msdn.microsoft.com/en-us/library/ff960302)

## Can I just include the source of this library in my project?

Yes! This is exactly how MFCMAPI does it. Read more about [Directly Including the MAPI Stub Library](DirectlyIncluding.md).

## Where are the instructions for building this library?

See: [Building the MAPI Stub Library](Building.md).

## What about instructions for linking to it?

See: [Linking To MAPIStubLibrary.lib](Linking.md).

## Why do I have to build this? Why can't I just download a pre-built .lib?

By providing full source and build projects for this library, we're giving complete control to the developer who uses it. This way, you're not dependent on us to provide new builds to correct any issues which may be uncovered in the library. Also, this gives the developer the flexibility to add and remove features from the library to fit their specific needs. This release mechanism does build in an assumption that the developer using this library is comfortable building a C++ project, but that's a safe assumption if they're already using MAPI.

## I use a compiler other than Visual Studio, or a different version of Visual Studio than one of the provided projects. What can I do?

You should still be able to use the provided source. You'll just have to build your own project or directly include the source in your existing project. Feel free to share your experiences in the [Discussions](https://github.com/microsoft/MAPIStubLibrary/discussions).

## I don't like the way the library does X. Can I change it?

Of course. Do whatever you like to the code. If you find bugs you wish to report, please visit the [Issues](https://github.com/microsoft/MAPIStubLibrary/issues) page.

## The MAPI Stub Library doesn't export the function X. Can I add it?

Yes. There are plenty of examples in [`MapiStubLibrary.cpp`](../library/mapiStubLibrary.cpp) of how to export a function. If the function is one we've documented, please file a "Missing Export" bug under Issues so we can consider adding it for everyone.

## The MAPI Stub Library exports the function X incorrectly. Can I fix it?

Yes. You can correct the export in [`MapiStubLibrary.cpp`](../library/mapiStubLibrary.cpp). Please also file an "Incorrect Export" bug under Issues so we can consider fixing it for everyone.