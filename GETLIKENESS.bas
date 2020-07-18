Public Function GETMAXLIKENESS(Arg1 As String, Arg2 As Variant) As Double
    Dim max As Double
    max = 0
    
    Dim foundOnce As Boolean
    foundOnce = False
    
    For Each cell In Arg2
        If cell = Arg1 And foundOnce = False Then
            foundOnce = True
        Else
            Dim likeness As Double
            likeness = Similarity(Arg1, cell)
            
            If max < likeness Then
                max = likeness
            End If
        End If
    Next cell
    
    
    GETMAXLIKENESS = max
End Function

Public Function GETLIKENESS(Arg1 As String, Arg2 As String) As Double
    GETLIKENESS = Similarity(Arg1, Arg2)
End Function

Private Function Similarity(ByVal String1 As String, ByVal String2 As String, Optional ByRef RetMatch As String, Optional min_match = 1) As Single
    Dim b1() As Byte, b2() As Byte
    Dim lngLen1 As Long, lngLen2 As Long
    Dim lngResult As Long
    
    If UCase(String1) = UCase(String2) Then 'Exactly the same
        Similarity = 1
    Else 
        lngLen1 = Len(String1)
        lngLen2 = Len(String2)
        If (lngLen1 = 0) Or (lngLen2 = 0) Then
            Similarity = 0 'Length of string is 0, return 0
        Else 'otherwise find similarity
            b1() = StrConv(UCase(String1), vbFromUnicode)
            b2() = StrConv(UCase(String2), vbFromUnicode)
            lngResult = Similarity_sub(0, lngLen1 - 1, 0, lngLen2 - 1, b1, b2, String1, RetMatch, min_match)
            Erase b1
            Erase b2
            If lngLen1 >= lngLen2 Then
                Similarity = lngResult / lngLen1
            Else
                Similarity = lngResult / lngLen2
            End If
        End If
    End If
 
End Function 
Private Function Similarity_sub(ByVal start1 As Long, ByVal end1 As Long,  ByVal start2 As Long, ByVal end2 As Long, _
                                ByRef b1() As Byte, ByRef b2() As Byte, ByVal FirstString As String, ByRef RetMatch As String, _
                                ByVal min_match As Long, Optional recur_level As Integer = 0) As Long
    Dim lngCurr1 As Long, lngCurr2 As Long
    Dim lngMatchAt1 As Long, lngMatchAt2 As Long
    Dim i As Long
    Dim lngLongestMatch As Long, lngLocalLongestMatch As Long
    Dim strRetMatch1 As String, strRetMatch2 As String
    
    If (start1 > end1) Or (start1 < 0) Or (end1 - start1 + 1 < min_match) Or (start2 > end2) Or (start2 < 0) Or (end2 - start2 + 1 < min_match) Then
        Exit Function '(exit if start/end is out of string, or length is too short)
    End If
    
    For lngCurr1 = start1 To end1 'loop through characters of first string
        For lngCurr2 = start2 To end2 'loop through characters of second string
        i = 0
        Do Until b1(lngCurr1 + i) <> b2(lngCurr2 + i) 'as long as characters match
            i = i + 1
            If i > lngLongestMatch Then 'if longer than previous, store starts & length
                lngMatchAt1 = lngCurr1
                lngMatchAt2 = lngCurr2
                lngLongestMatch = i
            End If
            If (lngCurr1 + i) > end1 Or (lngCurr2 + i) > end2 Then Exit Do
        Loop
        Next lngCurr2
    Next lngCurr1
    
    If lngLongestMatch < min_match Then Exit Function 'no matches at all, so no point checking for sub-matches!    
    lngLocalLongestMatch = lngLongestMatch  'call again for BEFORE + AFTER
    RetMatch = ""
    
    'Find longest match BEFORE the current position
    lngLongestMatch = lngLongestMatch + Similarity_sub(start1, lngMatchAt1 - 1, start2, lngMatchAt2 - 1, b1, b2, FirstString, strRetMatch1, min_match, recur_level + 1)
    If strRetMatch1 <> "" Then
        RetMatch = RetMatch & strRetMatch1 & "*"
    Else
        RetMatch = RetMatch & IIf(recur_level = 0 And lngLocalLongestMatch > 0 And (lngMatchAt1 > 1 Or lngMatchAt2 > 1) , "*", "")
    End If
    
    'add local longest
    RetMatch = RetMatch & Mid$(FirstString, lngMatchAt1 + 1, lngLocalLongestMatch)
                                
    'Find longest match AFTER the current position
    lngLongestMatch = lngLongestMatch + Similarity_sub(lngMatchAt1 + lngLocalLongestMatch, end1, lngMatchAt2 + lngLocalLongestMatch, end2, b1, b2, FirstString, strRetMatch2, min_match, recur_level + 1)    
    If strRetMatch2 <> "" Then
        RetMatch = RetMatch & "*" & strRetMatch2
    Else
        RetMatch = RetMatch & IIf(recur_level = 0 And lngLocalLongestMatch > 0 And ((lngMatchAt1 + lngLocalLongestMatch < end1) Or (lngMatchAt2 + lngLocalLongestMatch < end2)) , "*", "")
    End If
    
    
    'Return result
    Similarity_sub = lngLongestMatch
 
End Function
