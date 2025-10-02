# Building the MAPI Stub Library

To build the MAPI Stub Library as MAPIStubLibrary.lib, do the following:

## Getting the Source

You have two options to get the source code:

1. **Clone the repository (recommended)**:

   ```bash
   git clone https://github.com/microsoft/MAPIStubLibrary.git
   cd MAPIStubLibrary
   ```

2. **Download as ZIP**:
   Download the [latest source](https://github.com/microsoft/MAPIStubLibrary/archive/refs/heads/main.zip) and extract it to a local directory.

## Building with Visual Studio

The project includes all necessary MAPI headers in the `include` directory, so no additional downloads are required.

1. **Open the solution**: Open `mapistub.sln` in Visual Studio (2019 or later recommended).

2. **Select configuration**: From the Solution Configuration dropdown, select:
   - `Release` for production builds
   - `Debug` for debugging purposes

3. **Select platform**: From the Solution Platform dropdown, select:
   - `x64` for 64-bit applications (recommended)
   - `Win32` for 32-bit applications
   - `ARM64` for ARM64 applications

4. **Build**: Press `Ctrl+Shift+B` or go to Build → Build Solution.

You now have a file, `MAPIStubLibrary.lib`, which you can link in to your project.

## Building with Node.js/node-gyp

Alternatively, you can build using Node.js and node-gyp (this creates a static library):

1. **Install Node.js**: Make sure you have [Node.js](https://nodejs.org/) installed.

2. **Install dependencies**:

   ```bash
   npm install
   ```

3. **Build**:

   ```bash
   npm run build
   ```

The output will be `MAPIStubLibrary.lib` in the `build/Release` directory.

## Build Output

### Visual Studio Build

After building with Visual Studio, you'll find:

- `MAPIStubLibrary.lib` - The static library for linking
- Debug symbols (if building in Debug configuration)

### Node.js Build

After building with node-gyp, you'll find:

- `MAPIStubLibrary.lib` - The static library for linking in `build/Release/`

Both build methods produce the same `MAPIStubLibrary.lib` static library that you can link into your C++ projects.
