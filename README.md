# VB6 JsonSerializer Class Module(.cls)
Easily serialize json using the TypeLib Information com library from Microsoft.

# Setup
Copy the `JsonSerializer` class module into your project.

Add a project reference to `TypeLib Information` from tlbinf32.dll.

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
Dim testObj as CustomClass
Dim jsonStr as string

Set testObj = new CustomClass
jsonStr = helper.ConvertToJson(testObj)`
```

You can optionally enable(default) or disable null values from being in the json via the IgnoreNulls property

`helper.IgnoreNulls = false`
