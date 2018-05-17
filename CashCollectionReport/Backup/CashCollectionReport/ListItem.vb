Option Strict On
Option Explicit On 

Public Class ListItem
    Private mTimeRange As String
    Private mID As Integer

    Public Sub New(ByVal strValue As String, ByVal intID As Integer)
        mTimeRange = strValue
        mID = intID
    End Sub

    Public Sub New()
        mTimeRange = ""
        mID = 0
    End Sub

    Property ID() As Integer
        Get
            Return mID
        End Get
        Set(ByVal Value As Integer)
            mID = Value
        End Set
    End Property

    Property Value() As String
        Get
            Return mTimeRange
        End Get
        Set(ByVal Value As String)
            mTimeRange = Value
        End Set
    End Property

    ' When a list box displays an item in its collection
    ' it calls the ToString method to get the value to
    ' display so we will need to override this method
    ' so our class will return mTimeRange.
    Public Overrides Function ToString() As String
        Return mTimeRange
    End Function
End Class
