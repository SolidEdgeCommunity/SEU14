Imports System.Runtime.InteropServices
Module Main
   
    Public Sub Main()
        ' Get the type from the Solid Edge ProgID
        objSEType = Type.GetTypeFromProgID("SolidEdge.Application")
        ' Get the type from the Revision Manager ProgID
        objRevManType = Type.GetTypeFromProgID("RevisionManager.Application")

        OleMessageFilter.Register()

        'Add your code here!




        OleMessageFilter.Revoke()

    End Sub
End Module
