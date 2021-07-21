Namespace Novacode
	Public Class CustomProperty
'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not allow class members with the same name:
		Private name_Renamed As String
'INSTANT VB NOTE: The variable value was renamed since Visual Basic does not allow class members with the same name:
		Private value_Renamed As Object
'INSTANT VB NOTE: The variable type was renamed since Visual Basic does not allow class members with the same name:
		Private type_Renamed As String

		''' <summary>
		''' The name of this CustomProperty.
		''' </summary>
		Public ReadOnly Property Name() As String
			Get
				Return name_Renamed
			End Get
		End Property

		''' <summary>
		''' The value of this CustomProperty.
		''' </summary>
		Public ReadOnly Property Value() As Object
			Get
				Return value_Renamed
			End Get
		End Property

		Friend ReadOnly Property Type() As String
			Get
				Return type_Renamed
			End Get
		End Property

		Friend Sub New(ByVal name As String, ByVal type As String, ByVal value As String)
			Dim realValue As Object
			Select Case type
				Case "lpwstr"
					realValue = value
					Exit Select

				Case "i4"
					realValue = Integer.Parse(value, System.Globalization.CultureInfo.InvariantCulture)
					Exit Select

				Case "r8"
					realValue = Double.Parse(value, System.Globalization.CultureInfo.InvariantCulture)
					Exit Select

				Case "filetime"
					realValue = Date.Parse(value, System.Globalization.CultureInfo.InvariantCulture)
					Exit Select

				Case "bool"
					realValue = Boolean.Parse(value)
					Exit Select

				Case Else
					Throw New Exception()
			End Select

			Me.name_Renamed = name
			Me.type_Renamed = type
			Me.value_Renamed = realValue
		End Sub

		Private Sub New(ByVal name As String, ByVal type As String, ByVal value As Object)
			Me.name_Renamed = name
			Me.type_Renamed = type
			Me.value_Renamed = value
		End Sub

		''' <summary>
		''' Create a new CustomProperty to hold a string.
		''' </summary>
		''' <param name="name">The name of this CustomProperty.</param>
		''' <param name="value">The value of this CustomProperty.</param>
		Public Sub New(ByVal name As String, ByVal value As String)
			Me.New(name, "lpwstr", TryCast(value, Object))
		End Sub


		''' <summary>
		''' Create a new CustomProperty to hold an int.
		''' </summary>
		''' <param name="name">The name of this CustomProperty.</param>
		''' <param name="value">The value of this CustomProperty.</param>
		Public Sub New(ByVal name As String, ByVal value As Integer)
			Me.New(name, "i4", TryCast(value, Object))
		End Sub


		''' <summary>
		''' Create a new CustomProperty to hold a double.
		''' </summary>
		''' <param name="name">The name of this CustomProperty.</param>
		''' <param name="value">The value of this CustomProperty.</param>
		Public Sub New(ByVal name As String, ByVal value As Double)
			Me.New(name, "r8", TryCast(value, Object))
		End Sub


		''' <summary>
		''' Create a new CustomProperty to hold a DateTime.
		''' </summary>
		''' <param name="name">The name of this CustomProperty.</param>
		''' <param name="value">The value of this CustomProperty.</param>
		Public Sub New(ByVal name As String, ByVal value As Date)
			Me.New(name, "filetime", TryCast(value.ToUniversalTime(), Object))
		End Sub

		''' <summary>
		''' Create a new CustomProperty to hold a bool.
		''' </summary>
		''' <param name="name">The name of this CustomProperty.</param>
		''' <param name="value">The value of this CustomProperty.</param>
		Public Sub New(ByVal name As String, ByVal value As Boolean)
			Me.New(name, "bool", TryCast(value, Object))
		End Sub
	End Class
End Namespace
