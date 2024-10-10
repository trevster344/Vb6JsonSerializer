# VB6 JsonSerializer Class Module(.cls)
Easily serialize json using the TypeLib Information com library from Microsoft.

# Setup
Copy the `JsonSerializer` class module into your project.

Add a project reference to `TypeLib Information` from tlbinf32.dll.
Add a project reference to `Microsoft Scripting Runtime`

# Install Missing DLL
Download the tlbinf32.dll
Make sure you run command prompt as administrator and register the dll for com interop.
`C:\Windows\System32 regsvr32 tlbinf32.dll`
Requires a restart of VB6 to see.

Define a class module with the json properties you're expecting to use
Initialize an instance of it or an array and pass it into the `ConvertToJson` method

# Example:

```
Dim helper as new JsonSerializer
Dim testObj as Class1
Dim jsonStr as string

Set testObj = new Class1
jsonStr = helper.ConvertToJson(testObj)`
```
Result:

`{ "name": "Test","test Spaced Name": null }`

You can optionally enable(default) or disable null values from being in the json via the IgnoreNulls property

`helper.IgnoreNulls = false`

`{ "name": "Test" }`

# Nested Object Handling
The serializer can handle nesting. There is no checks for infinite loops so be careful.

```
Dim a As New Class1

Dim b(1 To 2) As New Class1
Set b(1) = New Class1
Set b(2) = New Class1

Dim c As New Class1
Dim d As New Class1
    
c.nested = d
b(2).nested = c
a.nested = b

jsonStr = helper.ConvertToJson(a)
```

`{"name": "Test","test Spaced Name": "","nested": [{"name": "Test","test Spaced Name": "","nested": null},{"name": "Test","test Spaced Name": "","nested": {"name": "Test","test Spaced Name": "","nested": {"name": "Test","test Spaced Name": "","nested": null}}}]}`
