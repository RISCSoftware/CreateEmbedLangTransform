# CreateEmbedLangTransform

CreateEmbedLangTransform is a tool for post-processing Windows Installer packages to create a single multi-lingual package from multiple individual packages, as they are created by the [WiX Toolset](https://wixtoolset.org/).

This could also be done using a combination of a few scripts from the Windows Installer SDK. This combines that task into a single C++ command line tool using the Windows Installer database API.

## Usage

The tool takes 3 command line arguments:

```bat
CreateEmbedLangTransform.exe <target-and-reference-msi> <source-msi> <language-identifier>
```

Example:

To embed the German translation into an English installer, call the tool like this:

```bat
CreateEmbedLangTransform.exe Setup.msi Setup_de-de.msi 1031
```

For a WiX MSBuild project (.wixproj), the following can be added as post-build event:

```bat
rem Copy the English installer into a language-neutral folder. It will become the multilingual one.
copy en-us\$(TargetFileName) .

rem Create and embed the German transform into the final installer.
CreateEmbedLangTransform.exe $(TargetFileName) de-de\$(TargetFileName) 1031

rem Other languages can be added by multiple invocations of CreateEmbedLangTransform.
```

## What this does

1. The tool creates a Windows Installer transform as the difference between the first and second argument and stores it as a temporary file (equivalent to the `WiGenXfm.vbs` script from the Windows Installer SDK).
1. It embeds the transform as a substorage under the name equal to the language ID (the third argument) (equivalent to the `WiSubStg.vbs` script).
1. To support WiX v.4, the tool also adds the language ID to the Summary information property "Template" if it isn't already, so that the MSI package is recognized as multilingual by Windows Installer (equivalent to the `WiLangId.vbs` script).
1. The temporary file created in step 1 is deleted again.

**Note:** This is an *undocumented* feature of Windows Installer. It is therefore not officially supported by Microsoft, and not supported by WiX!

## Building

The tool is currently built with Visual Studio 2022. It requires the "C++ Desktop" workload. It can probably be built with older versions as long as C++17 is supported.
