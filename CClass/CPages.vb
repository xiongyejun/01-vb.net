''' <summary>
''' ��ҳ�鿴����
''' </summary>
''' <remarks></remarks>
Class CPages
    ''' <summary>
    ''' ��ǰҳ��
    ''' </summary>
    ''' <remarks></remarks>
    Private Page_Index As Integer
    ''' <summary>
    ''' ��ҳ��
    ''' </summary>
    ''' <remarks></remarks>
    Private Page_Count As Integer
    ''' <summary>
    ''' ÿһҳ������
    ''' </summary>
    ''' <remarks></remarks>
    Private Page_Nums As Integer

    'Property PageIndex
    '    Get
    '        Return Page_Index
    '    End Get
    '    Set(ByVal value)
    '        Page_Index = value
    '    End Set
    'End Property
    ''' <summary>
    ''' ��һҳ�������0ҳ��ʱ�򷵻�false
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Pre() As Boolean
        If Page_Index Then
            Page_Index -= 1
            If Page_Index Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' ��һҳ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function NextP() As Boolean
        If Page_Index >= Page_Count - 1 Then
            Return False
        Else
            Page_Index += 1
            If Page_Index >= Page_Count - 1 Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Function First() As Boolean
        Page_Index = 0

        Return True
    End Function

    Function Last() As Boolean
        Page_Index = Page_Count - 1

        Return True
    End Function

    Function GetLimitOffset() As String
        Return String.Format(" limit {0} offset {1}", Page_Nums, (Page_Nums * Page_Index))
    End Function

    Sub New(ByVal Rows As Integer, ByVal PageNums As Integer, Optional ByVal PageIndex As Integer = 0)
        Page_Index = PageIndex
        Page_Nums = PageNums
        Page_Count = Rows \ PageNums
        If Rows Mod PageNums Then
            '����������Ҫ��һҳ
            Page_Count += 1
        End If
    End Sub

    Protected Overrides Sub Finalize()

    End Sub

End Class