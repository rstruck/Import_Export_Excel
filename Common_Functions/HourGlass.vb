

''using System;
''using System.Windows.Forms;

''public class HourGlass : IDisposable {
''  public HourGlass() {
''    Enabled = true;
''  }
''  public void Dispose() {
''    Enabled = false;
''  }
''  public static bool Enabled {
''    get { return Application.UseWaitCursor; }
''    set {
''      if (value == Application.UseWaitCursor) return;
''      Application.UseWaitCursor = value;
''      Form f = Form.ActiveForm;
''      if (f != null && f.Handle != null)   // Send WM_SETCURSOR
''        SendMessage(f.Handle, 0x20, f.Handle, (IntPtr)1);
''    }
''  }
''  [System.Runtime.InteropServices.DllImport("user32.dll")]
''  private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
''}
''End Class

''You can use it either directly by assigning HourGlass.Enabled or like this:

''    private void button1_Click(object sender, EventArgs e) {
''      using (new HourGlass()) {
''        // Do something that takes time...
''        System.Threading.Thread.Sleep(2000);
''      }
''    }
'Public Class Form_Operations
'    Implements IDisposable
'    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
'    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
'    End Function


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