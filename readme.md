# Excel C# Programming Test
Generate C# class with interface in single file

## Basics

1) Open admin shell prompt 
    - `Start-Process Powershell -Verb runAs`

2) Export dll
    - `csc /target:library .\cSharpTestLibrary.cs`

3) Register dll and create tlb file 
    - `regasm /codebase /tlb .\cSharpTestLibrary.dll`

4) Open up `Book.xlsm`

5) Open up the VBA view

6) Add reference to vba project via Tools -> References
    - use the features in the new reference created