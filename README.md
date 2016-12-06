# XrmToolBox.CloneFieldDefinitions
A Plugin for the XrmToolBox that allows to copy field definitions from one entity to another.
It supports most field types, even lookups.

# Known issues
- Copying of Customer Lookups (2016 feature) is not yet possible
- Source Fields can only be copied once, afterwards a restart is needed. This includes successful aswell as unsuccessful copy tries. This will be fixed soon, the reason for this behaviour is that the retrieved source data is edited directly, therefore the fields cannot be found on subsequential runs

# NuGet
This plugin is available as Nuget Package.

[![NuGet Badge](https://buildstats.info/nuget/MsDyn.Contrib.CloneFieldDefinitions)](https://www.nuget.org/packages/MsDyn.Contrib.CloneFieldDefinitions)

# Contributors
I very much appreciate Pull Requests to this project, therefore I'd like to mention everyone contributing to this project.

- Remon Kamel ([@harryremon](https://github.com/harryremon))

# Build
[![Build status](https://ci.appveyor.com/api/projects/status/a4keqi1hpwj2b73f/branch/master?svg=true)](https://ci.appveyor.com/project/DigitalFlow/xrmtoolbox-clonefielddefinitions/branch/master)
