Imports System
Imports System.Reflection
Imports System.Collections.Generic

' カスタムアトリビュートの定義
<AttributeUsage(AttributeTargets.Property)>
Public Class DictionaryKeyAttribute
    Inherits Attribute

    Public ReadOnly Property Key As String

    Public Sub New(ByVal key As String)
        Me.Key = key
    End Sub
End Class

Public Module DictionaryExtensions
    ' オブジェクトのプロパティを連想配列に変換するメソッド
    <System.Runtime.CompilerServices.Extension> 
    Public Function ToDictionary(ByVal obj As Object) As Dictionary(Of String, Object)
        Dim dict As New Dictionary(Of String, Object)
        Dim properties As PropertyInfo() = obj.GetType().GetProperties()

        For Each prop As PropertyInfo In properties
            Dim attribute As DictionaryKeyAttribute = prop.GetCustomAttribute(Of DictionaryKeyAttribute)()
            If attribute IsNot Nothing Then
                dict.Add(attribute.Key, prop.GetValue(obj, Nothing))
            End If
        Next

        Return dict
    End Function
End Module

Public Class MyData
    <DictionaryKey("id")>
    Public Property Id As Integer

    <DictionaryKey("name")>
    Public Property Name As String

    <DictionaryKey("desc")>
    Public Property Description As String
End Class

Module Program
    Sub Main(args As String())
        Console.WriteLine("Hello World!")

        Dim myData As New MyData With {.Id = 1, .Name = "Cross Tower", .Description = "pow pow!!"}
        Dim dictionary As Dictionary(Of String, Object) = myData.ToDictionary()

        For Each item As KeyValuePair(Of String, Object) In dictionary
            Console.WriteLine($"{item.Key}: {item.Value}")
        Next
    End Sub
End Module
