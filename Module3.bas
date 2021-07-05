Attribute VB_Name = "Module3"
Option Explicit
Public tokens As String
Public token() As String
Public source As String
Public Sub Lexer()
Dim i, cur, token_counter, n As Integer
Dim hold, keywords, char As String
Dim in_string As Boolean
Dim lines() As String

'the symbols that determine end of token
keywords = " ()=+-*\%^><!¡:." & Chr(13) & Chr(10) & Chr(9) & Chr(34)
'this counter to count the number of token
token_counter = 0
'get source program
'source = frmMain.ActiveForm.rtfText.Text
'split the source into lines
lines() = Split(source, vbNewLine)
'the variable that will hold tokens
tokens = ""
'variable for holding the currently scaned token
hold = ""
'flag to determine weather we are inside a string or not, because if inside string we must read all string what ever it contain
in_string = False

'main scaning loop
For i = 0 To UBound(lines)
   'add the line number to the tokens
   tokens = tokens & Chr(9) & " " ' "- " & i + 1 & " -" & vbNewLine
   token_counter = token_counter + 1
   'loop to scan inside one line
   For cur = 1 To Len(lines(i)) + 1 'we add + 1 because len function determine the length of the line without vbnewline
       'get a char
       char = Mid(lines(i), cur, 1)
       'if comment ignore the rest of the line
       If char = "#" And Not in_string Then
          tokens = tokens & hold & " "
          token_counter = token_counter + 1
          hold = ""
          Exit For
       End If
       'if the char not one of keywords add it to hold
       If (InStr(keywords, char) = 0 Or in_string) And char <> Chr(34) Then
          hold = hold & char
       Else
          'if we face a keyword so the token is finished so add it to tokens
          If Len(hold) Then
             tokens = tokens & hold & " "
             token_counter = token_counter + 1
             hold = ""
          End If
          'then we have to add the symbol also to tokens
          Select Case char
              ' if we face a space or newline or tab ignore
              Case " ", Chr(13), Chr(10), Chr(9), vbNewLine, ""
              'don't do anything
              
              'if one of the following determine weather they r with = or not
              Case "+", "-", "*", "\", "^", "%", "!", "="
                   If Mid(lines(i), cur + 1, 1) = "=" Then
                      cur = cur + 1
                      tokens = tokens & char & "=" & " "
                      token_counter = token_counter + 1
                   Else
                      tokens = tokens & char & " "
                      token_counter = token_counter + 1
                   End If
              'determine if >> or >= or > only
              Case ">"
                   If Mid(lines(i), cur + 1, 1) = ">" Then
                      cur = cur + 1
                      tokens = tokens & char & ">" & " "
                      token_counter = token_counter + 1
                   ElseIf Mid(lines(i), cur + 1, 1) = "=" Then
                      cur = cur + 1
                      tokens = tokens & char & "=" & " "
                      token_counter = token_counter + 1
                   Else
                      tokens = tokens & char & " "
                      token_counter = token_counter + 1
                   End If
              Case "<"
                   If Mid(lines(i), cur + 1, 1) = "<" Then
                      cur = cur + 1
                      tokens = tokens & char & "<" & " "
                      token_counter = token_counter + 1
                   ElseIf Mid(lines(i), cur + 1, 1) = "=" Then
                      cur = cur + 1
                      tokens = tokens & char & "=" & " "
                      token_counter = token_counter + 1
                   Else
                      tokens = tokens & char & " "
                      token_counter = token_counter + 1
                   End If
              'if we face a string begining or ending set or reset in_string flag
              Case Chr(34)
                   tokens = tokens & char & " "
                   token_counter = token_counter + 1
                   If in_string Then
                      in_string = False
                   Else
                      in_string = True
                   End If
              'if any other char from keywords just add it to tokens
              Case Else
                   tokens = tokens & char & " "
                   token_counter = token_counter + 1
                   
          End Select
       End If
   Next cur
Next i

ReDim token(token_counter) As String
Form1.Text1.Text = Replace(tokens, Chr(9), "")
Form1.Text1.Text = Replace(Form1.Text1.Text, " ", vbNewLine)
Form1.Show

'this part is to split the tokens into array because split function doesn't deal with spaces inside stings
'token() = Split(tokens, " ")
in_string = False
cur = 1
If InStr(tokens, " ") > 0 Then
  For i = 0 To token_counter
    n = InStr(cur, tokens, " ")
    If n = 0 Then
      Exit For
    End If
    hold = Mid(tokens, cur, n - cur)
    If Not in_string Then
     token(i) = hold
    ElseIf hold = Chr(34) Then
      i = i + 1
      token(i) = hold
    Else
      token(i) = token(i) & " " & hold
      i = i - 1
    End If
    If hold = Chr(34) Then
      If in_string Then
        in_string = False
      Else
        in_string = True
      End If
    End If
    cur = n + 1
  Next i
End If
    
'For i = 0 To UBound(token)
'   Form1.Text2.Text = Form1.Text2.Text & token(i) & " " '& Str(Asc(token(i))) & vbNewLine
'Next i

End Sub


