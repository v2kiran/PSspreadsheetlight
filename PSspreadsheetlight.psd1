﻿#
# Module manifest for module 'PSspreadsheetlight'
#
# Generated by: Kiran Reddy
#
# Generated on: 4/2/2023
#

@{

    # Script module or binary module file associated with this manifest.
     RootModule = 'PSspreadsheetlight.psm1'

    # Version number of this module.
    ModuleVersion = '1.0'

    # ID used to uniquely identify this module
    GUID = '0bbc52d6-3fa5-4b72-a75a-4cec828cecb9'

    # Author of this module
    Author = 'Kiran Reddy'

    # Company or vendor of this module
    CompanyName = 'Kiran Reddy'

    # Copyright statement for this module
    Copyright = '(c) 2023 Kiran Reddy. All rights reserved.'

    # Description of the functionality provided by this module
     Description = 'Powershell Module for Excel Automation'

    # Minimum version of the Windows PowerShell engine required by this module
     PowerShellVersion = '3.0'

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of the .NET Framework required by this module
     #DotNetFrameworkVersion = '3.5'

    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()

    # Assemblies that must be loaded prior to importing this module
     RequiredAssemblies = @('lib/SpreadsheetLight.dll','lib/DocumentFormat.OpenXml.dll','lib/System.Drawing.Common.dll','lib/NineDigit.SpreadSheetLightExtensions.dll')

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
     #ScriptsToProcess = @('')

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # Functions to export from this module
    FunctionsToExport = '*'

    # Cmdlets to export from this module
    #CmdletsToExport = '*sl*'

    # Variables to export from this module
    VariablesToExport = '*'

    # Aliases to export from this module
    AliasesToExport = '*'

    # List of all modules packaged with this module.
    ModuleList = @('PSspreadsheetlight.psm1')

    # List of all files packaged with this module
     FileList = @('PSspreadsheetlight.psm1')

    # Private data to pass to the module specified in RootModule/ModuleToProcess
    # PrivateData = ''

    # HelpInfo URI of this module
    # HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''

    }