Imports Microsoft.Win32

Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports VB6 = Microsoft.VisualBasic

Public Class ExtendedTreeNode
    Inherits TreeNode

    Property Key As Key

    Public Sub New(ByVal name As String)
        MyBase.New(name)
    End Sub
End Class

Public Class Key
    Property SubKeys As New List(Of Key)
    Property Name As String
    Property RootKeyNew As RegistryKey
    Property KeyPathNew As String

    Public Sub New(ByVal name As String, ByVal fullname As String)
        Me.Name = name
        Me.TreeNode = New ExtendedTreeNode(name)

        If fullname.StartsWith("HKEY_CLASSES_ROOT") Then
            RootKeyNew = Registry.ClassesRoot
        ElseIf fullname.StartsWith("HKEY_CURRENT_USER") Then
            RootKeyNew = Registry.CurrentUser
        ElseIf fullname.StartsWith("HKEY_LOCAL_MACHINE") Then
            RootKeyNew = Registry.LocalMachine
        ElseIf fullname.StartsWith("HKEY_USERS") Then
            RootKeyNew = Registry.Users
        ElseIf fullname.StartsWith("HKEY_CURRENT_CONFIG") Then
            RootKeyNew = Registry.CurrentConfig
        End If

        If fullname.Contains("\") Then
            KeyPathNew = Right(fullname, "\")
        End If
    End Sub

    Private TreeNodeValue As ExtendedTreeNode

    Public Property TreeNode() As ExtendedTreeNode
        Get
            Return TreeNodeValue
        End Get
        Set(ByVal Value As ExtendedTreeNode)
            TreeNodeValue = Value
            TreeNodeValue.Key = Me
        End Set
    End Property

    Public Sub BuildSubNoded()
        BuildSubNoded(Me)
    End Sub

    Public Shared Sub BuildSubNoded(ByVal k As Key)
        For Each i In k.GetSubKeyNames
            BuildSubNoded(k.GetSubKey(i))
        Next
    End Sub

    Public Function GetSubKey(ByVal name As String) As Key
        Using k = GetKey()
            If name.Contains("\") Then
                Dim tryKey = k.OpenSubKey(name)

                If Not tryKey Is Nothing Then
                    tryKey.Close()
                    Dim rootName = Left(name, "\")

                    Using firstkey = k.OpenSubKey(rootName)
                        If Not firstkey Is Nothing Then
                            For Each i In SubKeys
                                If i.Name = rootName Then
                                    Return i.GetSubKey(Right(name, "\"))
                                End If
                            Next

                            Dim ret As New Key(rootName, firstkey.Name)
                            TreeNode.Nodes.Add(ret.TreeNode)
                            SubKeys.Add(ret)
                            Return ret.GetSubKey(Right(name, "\"))
                        End If
                    End Using
                End If
            Else
                Using subkey = k.OpenSubKey(name)
                    If Not subkey Is Nothing AndAlso name <> "" Then
                        For Each i As Key In SubKeys
                            If i.Name = name Then
                                Return i
                            End If
                        Next

                        Dim ret As New Key(name, subkey.Name)
                        TreeNode.Nodes.Add(ret.TreeNode)
                        SubKeys.Add(ret)
                        Return ret
                    End If
                End Using
            End If

            Return Nothing
        End Using
    End Function

    Private Function GetKey() As RegistryKey
        If KeyPathNew Is Nothing Then
            Return RootKeyNew
        Else
            Return RootKeyNew.OpenSubKey(KeyPathNew)
        End If
    End Function

    Public Function GetValueNames() As List(Of String)
        Dim ret = New List(Of String)

        Using k = GetKey()
            For Each i In k.GetValueNames
                ret.Add(i)
            Next
        End Using

        Return ret
    End Function

    Public Function GetSubKeyNames() As List(Of String)
        Dim ret = New List(Of String)

        Using k = GetKey()
            For Each i In k.GetSubKeyNames
                ret.Add(i)
            Next
        End Using

        Return ret
    End Function

    Public Function GetValue(ByVal name As String) As Object
        Using k = GetKey()
            Return k.GetValue(name)
        End Using
    End Function

    Public Function GetNativeName() As String
        Using k = GetKey()
            Return k.Name
        End Using
    End Function
End Class

Module Misc
    Public Const CrLf As String = VB6.vbCrLf
    Public Const CrLf2 As String = VB6.vbCrLf + VB6.vbCrLf

    Public Function Left(ByVal value As String, ByVal start As String) As String
        If Not value.Contains(start) Then
            Return ""
        End If

        Return value.Substring(0, value.IndexOf(start))
    End Function

    Public Function Right(ByVal value As String, ByVal start As String) As String
        If Not value.Contains(start) Then
            Return ""
        End If

        Return value.Substring(value.IndexOf(start) + start.Length)
    End Function

    Public Function Msg(ByVal message As String, ByVal icon As MessageBoxIcon) As DialogResult
        Return MessageBox.Show(message, Application.ProductName, MessageBoxButtons.OK, icon)
    End Function
End Module

Public Class Native
    <DllImport("uxtheme.dll", CharSet:=CharSet.Unicode)>
    Shared Function SetWindowTheme(ByVal hWnd As IntPtr,
                                   ByVal pszSubAppName As String,
                                   ByVal pszSubIdList As String) As Integer
    End Function
End Class