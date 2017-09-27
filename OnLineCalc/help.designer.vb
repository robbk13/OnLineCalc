﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Linq
Imports System.Data.Linq.Mapping
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Reflection


<Global.System.Data.Linq.Mapping.DatabaseAttribute(Name:="lsp_calc")>  _
Partial Public Class helpDataContext
	Inherits System.Data.Linq.DataContext
	
	Private Shared mappingSource As System.Data.Linq.Mapping.MappingSource = New AttributeMappingSource()
	
  #Region "Extensibility Method Definitions"
  Partial Private Sub OnCreated()
  End Sub
  Partial Private Sub InsertContent(instance As Content)
    End Sub
  Partial Private Sub UpdateContent(instance As Content)
    End Sub
  Partial Private Sub DeleteContent(instance As Content)
    End Sub
  #End Region
	
	Public Sub New()
		MyBase.New(Global.System.Configuration.ConfigurationManager.ConnectionStrings("lsp_calcConnectionString1").ConnectionString, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As String)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As System.Data.IDbConnection)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As String, ByVal mappingSource As System.Data.Linq.Mapping.MappingSource)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As System.Data.IDbConnection, ByVal mappingSource As System.Data.Linq.Mapping.MappingSource)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public ReadOnly Property Contents() As System.Data.Linq.Table(Of Content)
		Get
			Return Me.GetTable(Of Content)
		End Get
	End Property
End Class

<Global.System.Data.Linq.Mapping.TableAttribute(Name:="dbo.Content")>  _
Partial Public Class Content
	Implements System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
	
	Private Shared emptyChangingEventArgs As PropertyChangingEventArgs = New PropertyChangingEventArgs(String.Empty)
	
	Private _ID As Integer
	
	Private _PageName As String
	
	Private _PageType As String
	
	Private _Heading1 As String
	
	Private _Block As String
	
	Private _Seq As System.Nullable(Of Integer)
	
	Private _Other As String
	
	Private _ContentGroup As String
	
	Private _ContentSubGroup As String
	
	Private _ImageURL As String
	
	Private _ImageHeight As System.Nullable(Of Integer)
	
	Private _ImageWidth As System.Nullable(Of Integer)
	
    #Region "Extensibility Method Definitions"
    Partial Private Sub OnLoaded()
    End Sub
    Partial Private Sub OnValidate(action As System.Data.Linq.ChangeAction)
    End Sub
    Partial Private Sub OnCreated()
    End Sub
    Partial Private Sub OnIDChanging(value As Integer)
    End Sub
    Partial Private Sub OnIDChanged()
    End Sub
    Partial Private Sub OnPageNameChanging(value As String)
    End Sub
    Partial Private Sub OnPageNameChanged()
    End Sub
    Partial Private Sub OnPageTypeChanging(value As String)
    End Sub
    Partial Private Sub OnPageTypeChanged()
    End Sub
    Partial Private Sub OnHeading1Changing(value As String)
    End Sub
    Partial Private Sub OnHeading1Changed()
    End Sub
    Partial Private Sub OnBlockChanging(value As String)
    End Sub
    Partial Private Sub OnBlockChanged()
    End Sub
    Partial Private Sub OnSeqChanging(value As System.Nullable(Of Integer))
    End Sub
    Partial Private Sub OnSeqChanged()
    End Sub
    Partial Private Sub OnOtherChanging(value As String)
    End Sub
    Partial Private Sub OnOtherChanged()
    End Sub
    Partial Private Sub OnContentGroupChanging(value As String)
    End Sub
    Partial Private Sub OnContentGroupChanged()
    End Sub
    Partial Private Sub OnContentSubGroupChanging(value As String)
    End Sub
    Partial Private Sub OnContentSubGroupChanged()
    End Sub
    Partial Private Sub OnImageURLChanging(value As String)
    End Sub
    Partial Private Sub OnImageURLChanged()
    End Sub
    Partial Private Sub OnImageHeightChanging(value As System.Nullable(Of Integer))
    End Sub
    Partial Private Sub OnImageHeightChanged()
    End Sub
    Partial Private Sub OnImageWidthChanging(value As System.Nullable(Of Integer))
    End Sub
    Partial Private Sub OnImageWidthChanged()
    End Sub
    #End Region
	
	Public Sub New()
		MyBase.New
		OnCreated
	End Sub
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ID", AutoSync:=AutoSync.OnInsert, DbType:="Int NOT NULL IDENTITY", IsPrimaryKey:=true, IsDbGenerated:=true)>  _
	Public Property ID() As Integer
		Get
			Return Me._ID
		End Get
		Set
			If ((Me._ID = value)  _
						= false) Then
				Me.OnIDChanging(value)
				Me.SendPropertyChanging
				Me._ID = value
				Me.SendPropertyChanged("ID")
				Me.OnIDChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_PageName", DbType:="VarChar(100)")>  _
	Public Property PageName() As String
		Get
			Return Me._PageName
		End Get
		Set
			If (String.Equals(Me._PageName, value) = false) Then
				Me.OnPageNameChanging(value)
				Me.SendPropertyChanging
				Me._PageName = value
				Me.SendPropertyChanged("PageName")
				Me.OnPageNameChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_PageType", DbType:="VarChar(100)")>  _
	Public Property PageType() As String
		Get
			Return Me._PageType
		End Get
		Set
			If (String.Equals(Me._PageType, value) = false) Then
				Me.OnPageTypeChanging(value)
				Me.SendPropertyChanging
				Me._PageType = value
				Me.SendPropertyChanged("PageType")
				Me.OnPageTypeChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Heading1", DbType:="VarChar(500)")>  _
	Public Property Heading1() As String
		Get
			Return Me._Heading1
		End Get
		Set
			If (String.Equals(Me._Heading1, value) = false) Then
				Me.OnHeading1Changing(value)
				Me.SendPropertyChanging
				Me._Heading1 = value
				Me.SendPropertyChanged("Heading1")
				Me.OnHeading1Changed
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Block", DbType:="VarChar(MAX)")>  _
	Public Property Block() As String
		Get
			Return Me._Block
		End Get
		Set
			If (String.Equals(Me._Block, value) = false) Then
				Me.OnBlockChanging(value)
				Me.SendPropertyChanging
				Me._Block = value
				Me.SendPropertyChanged("Block")
				Me.OnBlockChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Seq", DbType:="Int")>  _
	Public Property Seq() As System.Nullable(Of Integer)
		Get
			Return Me._Seq
		End Get
		Set
			If (Me._Seq.Equals(value) = false) Then
				Me.OnSeqChanging(value)
				Me.SendPropertyChanging
				Me._Seq = value
				Me.SendPropertyChanged("Seq")
				Me.OnSeqChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Other", DbType:="VarChar(1000)")>  _
	Public Property Other() As String
		Get
			Return Me._Other
		End Get
		Set
			If (String.Equals(Me._Other, value) = false) Then
				Me.OnOtherChanging(value)
				Me.SendPropertyChanging
				Me._Other = value
				Me.SendPropertyChanged("Other")
				Me.OnOtherChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ContentGroup", DbType:="VarChar(50)")>  _
	Public Property ContentGroup() As String
		Get
			Return Me._ContentGroup
		End Get
		Set
			If (String.Equals(Me._ContentGroup, value) = false) Then
				Me.OnContentGroupChanging(value)
				Me.SendPropertyChanging
				Me._ContentGroup = value
				Me.SendPropertyChanged("ContentGroup")
				Me.OnContentGroupChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ContentSubGroup", DbType:="VarChar(50)")>  _
	Public Property ContentSubGroup() As String
		Get
			Return Me._ContentSubGroup
		End Get
		Set
			If (String.Equals(Me._ContentSubGroup, value) = false) Then
				Me.OnContentSubGroupChanging(value)
				Me.SendPropertyChanging
				Me._ContentSubGroup = value
				Me.SendPropertyChanged("ContentSubGroup")
				Me.OnContentSubGroupChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ImageURL", DbType:="VarChar(150)")>  _
	Public Property ImageURL() As String
		Get
			Return Me._ImageURL
		End Get
		Set
			If (String.Equals(Me._ImageURL, value) = false) Then
				Me.OnImageURLChanging(value)
				Me.SendPropertyChanging
				Me._ImageURL = value
				Me.SendPropertyChanged("ImageURL")
				Me.OnImageURLChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ImageHeight", DbType:="Int")>  _
	Public Property ImageHeight() As System.Nullable(Of Integer)
		Get
			Return Me._ImageHeight
		End Get
		Set
			If (Me._ImageHeight.Equals(value) = false) Then
				Me.OnImageHeightChanging(value)
				Me.SendPropertyChanging
				Me._ImageHeight = value
				Me.SendPropertyChanged("ImageHeight")
				Me.OnImageHeightChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ImageWidth", DbType:="Int")>  _
	Public Property ImageWidth() As System.Nullable(Of Integer)
		Get
			Return Me._ImageWidth
		End Get
		Set
			If (Me._ImageWidth.Equals(value) = false) Then
				Me.OnImageWidthChanging(value)
				Me.SendPropertyChanging
				Me._ImageWidth = value
				Me.SendPropertyChanged("ImageWidth")
				Me.OnImageWidthChanged
			End If
		End Set
	End Property
	
	Public Event PropertyChanging As PropertyChangingEventHandler Implements System.ComponentModel.INotifyPropertyChanging.PropertyChanging
	
	Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
	
	Protected Overridable Sub SendPropertyChanging()
		If ((Me.PropertyChangingEvent Is Nothing)  _
					= false) Then
			RaiseEvent PropertyChanging(Me, emptyChangingEventArgs)
		End If
	End Sub
	
	Protected Overridable Sub SendPropertyChanged(ByVal propertyName As [String])
		If ((Me.PropertyChangedEvent Is Nothing)  _
					= false) Then
			RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
		End If
	End Sub
End Class