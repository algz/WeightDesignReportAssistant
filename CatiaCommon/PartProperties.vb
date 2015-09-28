Imports System
Imports System.ComponentModel

Public Class PartProperties
    <Description("零部件ID")>
    Public Property id() As String

    <Description("零部件名称")>
    Public Property name() As String

    <Description("零部件父节点ID")>
    Public Property parentID() As String

    <Description("零部件密度")>
    Public Property density() As Single

    <Description("零部件体积")>
    Public Property dimension() As Single

    <Description("零部件质量")>
    Public Property quality() As Single

    <Description("零部件面积")>
    Public Property area() As Single

    <Description("零部件重心")>
    Public Property centerOfGravity() As Single()

    <Description("零部件惯性矩阵")>
    Public Property inertiaMatrix() As Single()

    Public Sub New()
        id = "1"
        name = "name1"
        parentID = "0"
        density = 0.1F
        dimension = 0.1F
        quality = 0.1F
        area = 0.1F
        centerOfGravity = New Single() {0.1F, 0.2F, 0.3F}
        inertiaMatrix = New Single() {0.1F, 0.2F, 0.3F, 0.4F, 0.5F, 0.6F, 0.7F, 0.8F, 0.9F}
    End Sub
End Class