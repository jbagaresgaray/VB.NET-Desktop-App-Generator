Imports System.Windows.Forms
Imports System.Data.Common
Public Class frmProperties
    Private _ConnectionProperties As DbConnectionStringBuilder

    Public Property ConnectionProperties() As DbConnectionStringBuilder
        Get
            Return _ConnectionProperties
        End Get
        Set(ByVal value As DbConnectionStringBuilder)
            _ConnectionProperties = value
        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub frmProperties_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MainPG.SelectedObject = _ConnectionProperties
    End Sub
End Class