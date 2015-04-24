'Public Class Form_Operations
'    Implements IDisposable

'    Private disposedValue As Boolean = False ' To detect redundant calls

'    Public Sub New()
'        Enabled = True
'    End Sub

'    Public Property Enabled() As Boolean
'        Get
'            Return Application.UseWaitCursor
'        End Get
'        Set(ByVal value As Boolean)
'            If value = Application.UseWaitCursor Then Return
'            Application.UseWaitCursor = value
'            Dim f As Form = Form.ActiveForm
'            If (Not (f Is Nothing)) AndAlso (Not (f.Handle = vbNull)) Then
'                SendMessage(f.Handle, &H20, f.Handle, 1)
'            End If
'        End Set
'    End Property

'    ' IDisposable
'    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
'        If Not Me.disposedValue Then
'            Enabled = False
'        End If
'        Me.disposedValue = True
'    End Sub

'#Region " IDisposable Support "
'    ' This code added by Visual Basic to correctly implement the disposable pattern.
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' Do not change this code.
'        ' Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'        Dispose(True)
'        GC.SuppressFinalize(Me)
'    End Sub
'#End Region

'End Class


