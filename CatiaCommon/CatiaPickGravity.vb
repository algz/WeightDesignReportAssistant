
Imports System.Xml
Imports INFITF
Imports MECMOD
Imports KnowledgewareTypeLib
Imports ProductStructureTypeLib
Imports PARTITF
Imports System.Math

Public Class CatiaPickGravity
    Public Catia As INFITF.Application
    Public Selection1 As INFITF.Selection
    Private doc1 As Object
    Private SelectedObjectpro As Object
    Dim Re As Object
    Dim Fathernumber As Integer
    Dim Childrennumber As Integer

    Public Property FatherText As String
    Public Property lstPartPropertis As List(Of PartProperties) = New List(Of PartProperties)

    Public Function GetCatiaModel() As String
        Try
            Catia = GetObject(, "CATIA.Application")
        Catch ex As Exception
            MsgBox("正在启动CATIA，请耐心等待")
            Catia = CreateObject("CATIA.Application")
            Catia.Visible = True
        End Try
        Try
            doc1 = Catia.ActiveDocument
            Selection1 = doc1.Selection
            Selection1.Clear()
            'MsgBox(TypeName(doc1))
        Catch ex As Exception
            MsgBox("您没有打开任何文档", 64, "提示")
            Return Nothing
        End Try
        '------------------------------------------------
        ' '3、检测Catia是否有装配文档，否则提示用户创建Product
        '------------------------------------------------
        If TypeName(doc1) <> "ProductDocument" Then
            MsgBox("请将上游模型加载至装配下再运行！", 64, "提示")
            Return Nothing
        End If
        If Catia.GetWorkbenchId <> "Assembly" Then
            Catia.StartWorkbench("Assembly")
        End If
        '------------------------------------------------
        ' '4、拾取Product
        '------------------------------------------------
        Dim InputObjectType(0) As Object

        Try
            InputObjectType(0) = "Product"
            Selection1.Clear()
            Dim strResult As String = Selection1.SelectElement3(InputObjectType, "Select Product", True, CATMultiSelectionMode.CATMultiSelTriggWhenUserValidatesSelection, True)
            If Selection1.Count > 0 Then
                SelectedObjectpro = Selection1.Item(1).Value
                FatherText = SelectedObjectpro.PartNumber
                Selection1.Clear()
            End If


        Catch ex As Exception
            'MsgBox("选择了错误的catia元素，请重新选择正确的元素", 48)
            MsgBox("CATIA出现异常，请重新拾取父节点", 48)
        End Try

        Return FatherText

    End Function

    Public Function GetGravityInfo() As List(Of PartProperties)
        lstPartPropertis.Clear()
        Fathernumber = 0
        Childrennumber = 1
        Re = FindProduct(FatherText)
        Analysis(Re)

        Return lstPartPropertis
    End Function

    Function FindProduct(ProductNumber)
        Dim oActiveDoc
        oActiveDoc = Catia.ActiveDocument
        If TypeName(oActiveDoc) = "ProductDocument" Then
            Dim oSelection
            oSelection = oActiveDoc.Selection
            oSelection.Search("CATProductSearch.Product,all")

            If oSelection.count > 0 Then
                On Error Resume Next
                Do While Err.Number = 0
                    Dim oProduct = oSelection.FindObject("CATIAProduct")
                    If Err.Number = 0 And oProduct.PartNumber = ProductNumber Then
                        FindProduct = oProduct
                        Exit Function
                    End If
                Loop
            End If
        End If

        If TypeName(oActiveDoc) = "PartDocument" Then
            'MsgBox oActiveDoc.name
            FindProduct = oActiveDoc.Part
            Exit Function
        End If

        FindProduct = Nothing
    End Function

    Sub Analysis(FatherProduct)
        '判断是否拾取
        If FatherText = Nothing Then
            Exit Sub
        End If

        Dim oInertia0 As AnyObject = FatherProduct.GetTechnologicalObject("Inertia")
        Dim mass As Double
        Dim Density As Double
        Dim Volume As Double
        Dim Area As Double
        Dim Gxyz(2)
        Dim IxyzMatrix(8)
        Dim Gx As Double
        Dim Gy As Double
        Dim Gz As Double
        Dim Ixx As Double
        Dim Ixy As Double
        Dim Ixz As Double
        Dim Iyy As Double
        Dim Iyx As Double
        Dim Iyz As Double
        Dim Izx As Double
        Dim Izy As Double
        Dim Izz As Double
        FatherProduct.DescriptionRef = Childrennumber
        mass = oInertia0.Mass
        Density = oInertia0.Density
        Volume = FatherProduct.Analyze.Volume
        Area = FatherProduct.Analyze.WetArea
        Volume = Volume / 1000000000
        Area = Area / 1000000
        oInertia0.GetCOGPosition(Gxyz)
        oInertia0.GetInertiaMatrix(IxyzMatrix)
        Gx = Gxyz(0)
        Gy = Gxyz(1)
        Gz = Gxyz(2)
        Ixx = IxyzMatrix(0)
        Ixy = IxyzMatrix(1)
        Ixz = IxyzMatrix(2)
        Iyx = IxyzMatrix(3)
        Iyy = IxyzMatrix(4)
        Iyz = IxyzMatrix(5)
        Izx = IxyzMatrix(6)
        Izy = IxyzMatrix(7)
        Izz = IxyzMatrix(8)


        Dim newpart As PartProperties = New PartProperties

        '只单一写到文本框了，你需要处理下
        newpart.name = FatherProduct.partnumber
        newpart.id = Childrennumber

        '错误处理
        'On Error Resume Next
        'newpart.parentID = FatherProduct.Parent.Parent.DescriptionRef '无父级节点
        'If Err.Number <> 0 Then
        '    newpart.parentID = Childrennumber - 1
        'Else
        '    newpart.parentID = FatherProduct.Parent.Parent.DescriptionRef
        'End If
        'On Error GoTo 0
        If newpart.id = 1 Then
            newpart.parentID = Childrennumber - 1
        Else
            newpart.parentID = FatherProduct.Parent.Parent.DescriptionRef
        End If

        'newpart.density = Round(Density, 9)
        'newpart.dimension = Round(Volume, 9)
        'newpart.quality = Round(mass, 9)
        'newpart.area = Round(Area, 9)
        'newpart.centerOfGravity(0) = Round(Gx, 9)
        'newpart.centerOfGravity(1) = Round(Gy, 9)
        'newpart.centerOfGravity(2) = Round(Gz, 9)
        'newpart.inertiaMatrix(0) = Round(Ixx, 9)
        'newpart.inertiaMatrix(1) = Round(Ixy, 9)
        'newpart.inertiaMatrix(2) = Round(Ixz, 9)
        'newpart.inertiaMatrix(3) = Round(Iyx, 9)
        'newpart.inertiaMatrix(4) = Round(Iyy, 9)
        'newpart.inertiaMatrix(5) = Round(Iyz, 9)
        'newpart.inertiaMatrix(6) = Round(Izx, 9)
        'newpart.inertiaMatrix(7) = Round(Izy, 9)
        'newpart.inertiaMatrix(8) = Round(Izz, 9)

        newpart.density = Density
        newpart.dimension = Volume
        newpart.quality = mass
        newpart.area = Area
        newpart.centerOfGravity(0) = Gx
        newpart.centerOfGravity(1) = Gy
        newpart.centerOfGravity(2) = Gz
        newpart.inertiaMatrix(0) = Ixx
        newpart.inertiaMatrix(1) = Ixy
        newpart.inertiaMatrix(2) = Ixz
        newpart.inertiaMatrix(3) = Iyx
        newpart.inertiaMatrix(4) = Iyy
        newpart.inertiaMatrix(5) = Iyz
        newpart.inertiaMatrix(6) = Izx
        newpart.inertiaMatrix(7) = Izy
        newpart.inertiaMatrix(8) = Izz

        lstPartPropertis.Add(newpart)

        Dim ChildrenP
        ChildrenP = FatherProduct.products
        Dim oCurrentProduct
        'If FatherProduct.products.count = 0 Then
        '    Selection1 = Catia.ActiveDocument.Selection
        '    Selection1.Clear()
        '    Selection1.Add(FatherProduct)

        'End If


        If Not ChildrenP Is Nothing Then

            For Each oCurrentProduct In ChildrenP

                'Selection1 = Catia.ActiveDocument.Selection
                'Selection1.Clear()
                'Selection1.Add(oCurrentProduct)

                Dim newpart1 As PartProperties = New PartProperties

                Childrennumber = Childrennumber + 1
                oCurrentProduct.DescriptionRef = Childrennumber
                If oCurrentProduct.Products.Count = 0 Then

                    Dim oInertia As AnyObject = oCurrentProduct.GetTechnologicalObject("Inertia")
                    mass = oInertia.Mass
                    Density = oInertia.Density
                    Volume = oCurrentProduct.Analyze.Volume
                    Area = oCurrentProduct.Analyze.WetArea
                    Volume = Volume / 1000000000
                    Area = Area / 1000000
                    oInertia.GetCOGPosition(Gxyz)
                    oInertia.GetInertiaMatrix(IxyzMatrix)
                    Gx = Gxyz(0)
                    Gy = Gxyz(1)
                    Gz = Gxyz(2)
                    Ixx = IxyzMatrix(0)
                    Ixy = IxyzMatrix(1)
                    Ixz = IxyzMatrix(2)
                    Iyx = IxyzMatrix(3)
                    Iyy = IxyzMatrix(4)
                    Iyz = IxyzMatrix(5)
                    Izx = IxyzMatrix(6)
                    Izy = IxyzMatrix(7)
                    Izz = IxyzMatrix(8)
                    'Fathernumber = FatherProduct.DescriptionRef
                    newpart1.name = oCurrentProduct.partnumber
                    newpart1.id = Childrennumber
                    newpart1.parentID = oCurrentProduct.Parent.Parent.DescriptionRef


                    '只单一写到文本框了，你需要处理下
                    'newpart1.density = Round(Density, 9)
                    'newpart1.dimension = Round(Volume, 9)
                    'newpart1.quality = Round(mass, 9)

                    'newpart1.area = Round(Area, 9)
                    'newpart1.centerOfGravity(0) = Round(Gx, 9)
                    'newpart1.centerOfGravity(1) = Round(Gy, 9)
                    'newpart1.centerOfGravity(2) = Round(Gz, 9)
                    'newpart1.inertiaMatrix(0) = Round(Ixx, 9)
                    'newpart1.inertiaMatrix(1) = Round(Ixy, 9)
                    'newpart1.inertiaMatrix(2) = Round(Ixz, 9)
                    'newpart1.inertiaMatrix(3) = Round(Iyx, 9)
                    'newpart1.inertiaMatrix(4) = Round(Iyy, 9)
                    'newpart1.inertiaMatrix(5) = Round(Iyz, 9)
                    'newpart1.inertiaMatrix(6) = Round(Izx, 9)
                    'newpart1.inertiaMatrix(7) = Round(Izy, 9)
                    'newpart1.inertiaMatrix(8) = Round(Izz, 9)

                    newpart1.density = Density
                    newpart1.dimension = Volume
                    newpart1.quality = mass
                    newpart1.area = Area
                    newpart1.centerOfGravity(0) = Gx
                    newpart1.centerOfGravity(1) = Gy
                    newpart1.centerOfGravity(2) = Gz
                    newpart1.inertiaMatrix(0) = Ixx
                    newpart1.inertiaMatrix(1) = Ixy
                    newpart1.inertiaMatrix(2) = Ixz
                    newpart1.inertiaMatrix(3) = Iyx
                    newpart1.inertiaMatrix(4) = Iyy
                    newpart1.inertiaMatrix(5) = Iyz
                    newpart1.inertiaMatrix(6) = Izx
                    newpart1.inertiaMatrix(7) = Izy
                    newpart1.inertiaMatrix(8) = Izz

                    lstPartPropertis.Add(newpart1)
                Else

                    Call Analysis(oCurrentProduct)

                End If

            Next
        End If

    End Sub



End Class
