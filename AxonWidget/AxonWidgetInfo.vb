Imports Grasshopper.Kernel

Public Class AxonWidgetInfo
    Inherits GH_AssemblyInfo

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "AxonWidget"
        End Get
    End Property
    Public Overrides ReadOnly Property Icon As System.Drawing.Bitmap
        Get
            'Return a 24x24 pixel bitmap to represent this GHA library.
            Return Nothing
        End Get
    End Property
    Public Overrides ReadOnly Property Description As String
        Get
            'Return a short string describing the purpose of this GHA library.
            Return ""
        End Get
    End Property
    Public Overrides ReadOnly Property Id As System.Guid
        Get
            Return New System.Guid("12a4e725-ab60-4c50-be4b-b0702da5c560")
        End Get
    End Property

    Public Overrides ReadOnly Property AuthorName As String
        Get
            Return "Daniel Abalde"
        End Get
    End Property

    Public Overrides ReadOnly Property AuthorContact As String
        Get
            Return "dga_3@hotmail.com"
        End Get
    End Property
End Class
