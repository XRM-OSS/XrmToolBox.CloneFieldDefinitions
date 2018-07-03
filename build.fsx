// include Fake libs
#I @"tools\FAKE\tools\"
#r @"tools\FAKE\tools\FakeLib.dll"

open Fake
open Fake.AssemblyInfoFile
open System.IO

//Project config
let projectName = "MsDyn.Contrib.CloneFieldDefinitions"
let projectDescription = "A XrmToolBox Plugin that helps copying field definitions from an entity to another"
let authors = ["Florian Kroenert"]

// Directories
let buildDir  = @".\build\"
let libbuildDir = buildDir + @"lib\"

let testDir   = @".\test\"

let deployDir = @".\Publish\"
let libdeployDir = deployDir + @"lib\"
let nugetDir = @".\nuget\"
let packagesDir = @".\packages\"

// version info
let mutable majorversion    = "1"
let mutable minorversion    = "1"
let mutable build           = "0"
let mutable nugetVersion    = ""
let mutable asmVersion      = ""
let mutable asmInfoVersion  = ""

let sha                     = Git.Information.getCurrentHash() 

// Targets
Target "Clean" (fun _ ->

    CleanDirs [buildDir; testDir; deployDir; nugetDir]
)

Target "RestorePackages" (fun _ ->

   let RestorePackages2() =
     !! "./src/**/packages.config"
     |> Seq.iter ( RestorePackage (fun p -> {p with Sources = ["http://go.microsoft.com/fwlink/?LinkID=206669"] } ))
     ()

   RestorePackages2()
)

Target "AssemblyInfo" (fun _ ->
    BulkReplaceAssemblyInfoVersions "src" (fun f -> 
                                              {f with
                                                  AssemblyVersion = asmVersion
                                                  AssemblyInformationalVersion = asmInfoVersion
                                                  AssemblyFileVersion = asmVersion})
)

Target "BuildLib" (fun _ ->
    !! @"src\lib\**\*.csproj"
        |> MSBuildRelease libbuildDir "Build"
        |> Log "Build-Output: "
)

Target "Publish" (fun _ ->
    CreateDir libdeployDir

    !! (libbuildDir @@ @"MsDyn*.dll")
            |> CopyTo libdeployDir
)

Target "CreateNuget" (fun _ ->
    "MsDyn.Contrib.CloneFieldDefinitions.nuspec"
          |> NuGet (fun p ->
            {p with
                Authors = authors
                Project = projectName
                Version = nugetVersion
                NoPackageAnalysis = true
                Description = projectDescription
                ToolPath = @".\tools\Nuget\Nuget.exe"
                OutputPath = nugetDir })
)

Target "PublishNuget" (fun _ ->

  let nugetPublishDir = (deployDir + "nuget")
  CreateDir nugetPublishDir

  !! (nugetDir + "*.nupkg")
     |> Copy nugetPublishDir
)

// Dependencies
"Clean"
  ==> "RestorePackages"
  =?> ("AssemblyInfo", not isLocalBuild )
  ==> "BuildLib"
  ==> "Publish"
  ==> "CreateNuget"
  ==> "PublishNuget"

// start build
RunTargetOrDefault "Publish"
