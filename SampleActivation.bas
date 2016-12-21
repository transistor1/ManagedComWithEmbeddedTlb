Attribute VB_Name = "SampleActivation"

'https://github.com/transistor1/ManagedComWithEmbeddedTlb

'This is a sample module that can be imported into a VBA project
'Make sure to deploy your DLL in the same folder as your VBA project
'Also, make sure to add a reference to the .DLL using Tools->References.

'Once the lib has been loaded with LoadLibrary (see property ComObj), then VBA will automagically
'know how to use the CreateComObject reference below.  The trouble is initially finding the COM object, because
'VBA doesn't look in the current folder; it looks in CurDir.

#If VBA7 Then

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpFileName As String) As Long
Private Declare PtrSafe Function CreateComObject Lib "ComWithEmbeddedTypeLib.dll" Alias "CreateObject" () As Object

#Else

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpFileName As String) As Long
Private Declare Function CreateComObject Lib "ComWithEmbeddedTypeLib.dll" Alias "CreateObject" () As Object

#End If

Function TestComObjectActivation()
    'Run this function to test your COM object. Copy it in the same
    'folder where this database lives.
    
    'Lazy-load the COM object
    Debug.Print ComObj.HelloWorld("activated, managed COM object")
End Function

Public Property Get ThisFolder(ParamArray additionalPaths())
    Dim fs__
    Set fs__ = CreateObject("Scripting.FileSystemObject")
    ThisFolder = fs__.GetParentFolderName(CurrentDb().Name)
    
    For Each curPath In additionalPaths
        ThisFolder = fs__.BuildPath(ThisFolder, curPath)
    Next
End Property

Public Property Get ComObj() As ComLib.IComWithEmbeddedTypeLib
    Static comLib__ As ComLib.IComWithEmbeddedTypeLib
    If comLib__ Is Nothing Then
        Dim libPath As String
        Dim result As Long
        
        libPath = ThisFolder("ComWithEmbeddedTypeLib.dll")
        result = LoadLibrary(libPath)
        
        If result <> 0 Then
            Set comLib__ = CreateComObject
        End If
    End If

    Set ComObj = comLib__
End Property
