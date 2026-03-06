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

## Building with Node.js scripts

The npm scripts now use MSBuild by default (and keep node-gyp scripts under `gyp:*`):

1. **Install Node.js**: Make sure you have [Node.js](https://nodejs.org/) installed.

2. **Install dependencies**:

   ```bash
   npm install
   ```

3. **Build with MSBuild**:

   **Default build:**

   ```bash
   npm run build
   ```

   **Common build variants:**

   ```bash
   npm run build:debug:x64      # default script target
   npm run build:release:x64
   npm run build:debug:x86
   npm run build:release:x86
   npm run build:debug:arm64
   npm run build:release:arm64
   npm run build:debug:arm64ec
   npm run build:release:arm64ec
   npm run build:all            # x64 + x86 all variants
   ```

   **Clean build outputs:**

   ```bash
   npm run clean
   ```

4. **(Optional) Build with node-gyp (legacy):**

   ```bash
   npm run gyp:build
   npm run gyp:build:x64
   npm run gyp:build:x86
   npm run gyp:build:arm64
   npm run gyp:clean
   ```

MSBuild outputs go to Visual Studio configuration/platform output directories. node-gyp outputs are in architecture-specific directories:

- `build/lib/x64/MAPIStubLibrary.lib` - 64-bit library
- `build/lib/ia32/MAPIStubLibrary.lib` - 32-bit library  
- `build/lib/arm64/MAPIStubLibrary.lib` - ARM64 library

## Build Output

### Visual Studio Build

After building with Visual Studio, you'll find:

- `MAPIStubLibrary.lib` - The static library for linking
- Debug symbols (if building in Debug configuration)

### Node.js (node-gyp) Build

After building with node-gyp (`gyp:*` scripts), you'll find:

- `MAPIStubLibrary.lib` - The static library for linking in `build/Release/`

Both build methods produce the same `MAPIStubLibrary.lib` static library that you can link into your C++ projects.
