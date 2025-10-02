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
      "msvs_settings": {
        "VCCLCompilerTool": {
          "RuntimeLibrary": 0, # /MT - static runtime
          "ExceptionHandling": 0, # disable exceptions
          "OptimizeReferences": 2, # /OPT:REF
          "EnableCOMDATFolding": 2, # /OPT:ICF
          "Optimization": 2, # /O2 instead of aggressive /Ox
        },
        "VCLibrarianTool": {
          "AdditionalOptions": ["/LTCG"] # Link Time Code Generation
        }
      },
      "defines": [
        "WIN32_LEAN_AND_MEAN",
        "NOMINMAX"
      ],
      "conditions": [
        ["OS=='win'", {
          "msvs_settings": {
            "VCCLCompilerTool": {
              "WarningLevel": 4,
              "DisableSpecificWarnings": ["4996"] # disable deprecated warnings
            }
          }
        }]
      ]
    }
  ]
}