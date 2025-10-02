{
  "targets": [
    {
      "target_name": "MAPIStubLibrary",
      "type": "static_library",
      "include_dirs": ["include"],
      "sources": [
        "library/mapiStubLibrary.cpp",
        "library/stubutils.cpp"
      ],
      "msvs_version": "2022",
      "msvs_settings": {
        "VCCLCompilerTool": {
          "RuntimeLibrary": 0, # /MT - static runtime
          "ExceptionHandling": 0, # disable exceptions
          "Optimization": 2, # /O2 instead of aggressive /Ox
          "WarningLevel": 4
        },
        "VCLibrarianTool": {
          "AdditionalOptions": ["/LTCG"] # Link Time Code Generation
        }
      },
      "defines": [
        "WIN32_LEAN_AND_MEAN",
        "NOMINMAX"
      ]
    }
  ]
}