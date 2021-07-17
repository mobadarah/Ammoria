Attribute VB_Name = "Module2"
Option Explicit
Dim program_name, Integer_Variables, Fractional_Variables, String_Variables, Integer_Constants, Fractional_Constants, String_Constants, Labels, Variables, Constants As String
Dim i, line_no, var_start, paranthes_counter As Integer
'«·„ €Ì—«  «· «·Ì… ⁄»«—… ⁄‰ ⁄œ«œ«  ·—”„ «· »«—”  —Ì
Dim Program_c, Stmt_list_c, Stmt_c, Other_stmt_c, Dec_stmt_c, Eq_stmt_c, If_stmt_c, While_stmt_c, Do_stmt_c, For_stmt_c, Switch_stmt_c, Go_stmt_c, Print_stmt_c, Input_stmt_c, Label_stmt_c, _
    Const_dec_c, Var_dec_c, Const_name_c, Const_type_c, Intg_no_c, Fract_no_c, String_c, _
    Var_names_c, Var_type_c, Var_na_c, Other_var_c, Cond_c, If_rest_c, Expression_c, Dec_Inc_c, _
    Cases_list_c, Label_na_c, Eq_op_c, Sign_c, Phrase_c, Expre_rest_c, Arit_fact_c, Phrase_resr_c, _
    Fact_c, Fact_rest_c, Direct_cond_c, Cond_rest_c, Direct_cond_rest_c, Print_list_c, Print_list_rest_c, _
    Input_list_c, Input_list_rest_c, Scope As Integer
Dim Parsed_Tree As String
Dim else_f As Boolean

Public Sub Parser()
Program_c = 0
Stmt_list_c = 0
Stmt_c = 0
Other_stmt_c = 0
Dec_stmt_c = 0
Eq_stmt_c = 0
If_stmt_c = 0
While_stmt_c = 0
Do_stmt_c = 0
For_stmt_c = 0
Switch_stmt_c = 0
Go_stmt_c = 0
Print_stmt_c = 0
Input_stmt_c = 0
Label_stmt_c = 0
Const_dec_c = 0
Var_dec_c = 0
Const_name_c = 0
Const_type_c = 0
Intg_no_c = 0
Fract_no_c = 0
String_c = 0
Var_names_c = 0
Var_type_c = 0
Var_na_c = 0
Other_var_c = 0
Cond_c = 0
If_rest_c = 0
Expression_c = 0
Dec_Inc_c = 0
Cases_list_c = 0
Label_na_c = 0
Eq_op_c = 0
Sign_c = 0
Phrase_c = 0
Expre_rest_c = 0
Arit_fact_c = 0
Phrase_resr_c = 0
Fact_c = 0
Fact_rest_c = 0
Direct_cond_c = 0
Cond_rest_c = 0
Direct_cond_rest_c = 0
Print_list_c = 0
Print_list_rest_c = 0
Input_list_c = 0
Input_list_rest_c = 0

Parsed_Tree = ""
Scope = 0
paranthes_counter = 0
i = 1
line_no = 1
Program
Form4.Text1.Text = Parsed_Tree
Form4.Show
MsgBox "«·»—‰«„Ã œﬁﬁ »‰Ã«Õ"
 End Sub
Function Get_token() As String
If i <= UBound(token) Then
   If token(i) = Chr(9) Then
      line_no = line_no + 1
   End If
   Get_token = token(i)
   i = i + 1
Else
   Get_token = Chr(9)
End If
End Function
Function Add_Tabs(ByVal no_of_tabs As Integer, ByVal strOriginal As String) As String

Dim c As Integer
Dim m As String

m = ""
If no_of_tabs > 0 Then
  For c = 1 To no_of_tabs
    m = m & Chr(9)
  Next c

  m = m & strOriginal
  Add_Tabs = m
Else
  Add_Tabs = strOriginal
End If
End Function

Sub Program()
Parsed_Tree = Parsed_Tree & "<«·»—‰«„Ã>" & vbNewLine
While Get_token() = Chr(9) Or token(i - 1) = " " Or token(i - 1) = ""
Wend
If token(i - 1) = "»œ«Ì…" Then
   Scope = Scope + 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "»œ«Ì…" & vbNewLine
   If Get_token() = "(" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ «·»—‰«„Ã>" & vbNewLine
      Scope = Scope + 1
      program_name = Get_token()
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & program_name & vbNewLine
      Scope = Scope - 1
      If Get_token() = ")" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
        If Get_token() = Chr(9) Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "”ÿ— ÃœÌœ" & vbNewLine
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
          Scope = Scope + 1
          Statment_List
          Scope = Scope - 1
          If token(i - 1) = "‰Â«Ì…" Then
            Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
            If Get_token() = "(" Then
              Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
              Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ «·»—‰«„Ã>" & vbNewLine
              If Get_token() = program_name Then
                Scope = Scope + 1
                Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & program_name & vbNewLine
                Scope = Scope - 1
                If Get_token() = ")" Then
                  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
                  Scope = Scope - 1
                End If
              End If
            End If
          End If
        Else
          MsgBox "‰ Êﬁ⁄ „‰ﬂ ‰Â«Ì… «·”ÿ—"
        End If
      Else
         MsgBox "‰ Êﬁ⁄ „‰ﬂ )"
      End If
    Else
      MsgBox "‰ Êﬁ⁄ „‰ﬂ ("
    End If
Else
    MsgBox "‰ Êﬁ⁄ „‰ﬂ »œ«Ì…"
End If

End Sub
Sub Statment_List()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·…>" & vbNewLine
Scope = Scope + 1
Statment
Scope = Scope - 1
If Get_token() = Chr(9) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "”ÿ— ÃœÌœ" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„· ¬Œ—Ï>" & vbNewLine
  Scope = Scope + 1
  Rest_of_statments
  Scope = Scope - 1
Else
  MsgBox "‰ Êﬁ⁄ „‰ﬂ ‰Â«Ì… «·”ÿ—"
End If
End Sub
Sub Rest_of_statments()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·…>" & vbNewLine
Scope = Scope + 1
Statment
Scope = Scope - 1
If token(i - 1) = "‰Â«Ì…" Or token(i - 1) = "ÿ«·„«" Or else_f Then
   Exit Sub
End If
If Get_token() = Chr(9) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "”ÿ— ÃœÌœ" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„· ¬Œ—Ï>" & vbNewLine
  Scope = Scope + 1
  Rest_of_statments
  Scope = Scope - 1
Else
  MsgBox "‰ Êﬁ⁄ „‰ﬂ ‰Â«Ì… «·”ÿ—° ›Ì «·”ÿ— —ﬁ„ " & line_no
End If
End Sub
Sub Statment()
else_f = False '··œ·«·… ⁄·Ï "Ê €Ì— –·ﬂ "Õ Ï  Œ—Ã „‰ «·sub
If Get_token() = "‰Â«Ì…" Or token(i - 1) = "ÿ«·„«" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
   Exit Sub
ElseIf token(i - 1) = "⁄—¯›" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·…  ⁄—Ì›>" & vbNewLine
   Scope = Scope + 1
   Declaration_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "≈–«" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… ≈–«>" & vbNewLine
   Scope = Scope + 1
   if_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "»Ì‰„«" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… »Ì‰„«>" & vbNewLine
   Scope = Scope + 1
   while_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "«›⁄·" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… «›⁄·>" & vbNewLine
   Scope = Scope + 1
   do_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "„‰" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… „‰ ≈·Ï>" & vbNewLine
   Scope = Scope + 1
   for_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "÷⁄" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… ÷⁄ ··«Œ Ì«—>" & vbNewLine
   Scope = Scope + 1
   switch_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "«–Â»" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… «–Â» ≈·Ï>" & vbNewLine
   Scope = Scope + 1
   go_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "«ÿ»⁄" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… «ÿ»⁄>" & vbNewLine
   Scope = Scope + 1
   Print_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = "√œŒ·" Then
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… √œŒ·>" & vbNewLine
   Scope = Scope + 1
   Input_stamt
   Scope = Scope - 1
ElseIf token(i - 1) = Chr(9) Then
    '≈–« ’«œ›‰« ‰Â«Ì… ”ÿ— ⁄·Ì‰« «—Ã«⁄ «·„ƒ‘— Œ«‰… ··Œ·› Ê«‰ﬁ«’ ⁄œœ «·«”ÿ— ·√‰ «·Õ’Ê· ⁄·Ï  Êﬂ‰ ”Ê› Ì⁄„· ‰›” «·⁄„·
    i = i - 1
    line_no = line_no - 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "”ÿ— ÃœÌœ" & vbNewLine
    Exit Sub
ElseIf (i + 2 <= UBound(token)) Then
   If ((token(i - 1) & token(i) & token(i + 1)) Like "*=*") Then
     i = i - 1 '«—Ã⁄‰« «·„ƒ‘— Œ«‰… ··Œ·› ·√‰ Ã„·… «·≈”‰«œ ”Ê›  ﬁÊ„ »⁄„· get_token()
     Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… ≈”‰«œ>" & vbNewLine
     Scope = Scope + 1
     Equality_stamt
     Scope = Scope - 1
   ElseIf ((token(i - 1) & " " & token(i) & " " & token(i + 1)) = "Ê €Ì— –·ﬂ") Then
     else_f = True
     Exit Sub
   ElseIf (token(i - 1) & token(i)) Like "*:" Then
     Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<⁄‰Ê«‰>" & vbNewLine
     Scope = Scope + 1
     Label_stamt
     Scope = Scope - 1
   Else
     MsgBox "Ã„·… €Ì— „⁄—Ê›… ›Ì «·”ÿ— —ﬁ„" & line_no
   End If
Else
   MsgBox "Ã„·… €Ì— „⁄—Ê›… ›Ì «·”ÿ— —ﬁ„" & line_no
End If

End Sub
Sub Declaration_stamt()
If (i + 1 <= UBound(token)) Then
  If token(i + 1) = "ﬂÀ«» " Then
     Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·…  ⁄—Ì› À«» >" & vbNewLine
     Scope = Scope + 1
     Constant_Decleration
     Scope = Scope - 1
  Else
     Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·…  ⁄—Ì› „ €Ì—>" & vbNewLine
     Scope = Scope + 1
     Variable_Deleration
     Scope = Scope - 1
  End If
End If

     
End Sub
Sub Constant_Decleration()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "⁄—¯ˆ›" & vbNewLine
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ À«» >" & vbNewLine
Scope = Scope + 1
Constant_Name
Scope = Scope - 1
i = i + 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "ﬂÀ«» " & vbNewLine
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‰Ê⁄ «·À«» >" & vbNewLine
Scope = Scope + 1
Constant_Type
Scope = Scope - 1
End Sub
Sub Constant_Name()
If Not Get_token() Like "[¡-Ì]*" Then
   MsgBox "«”„ «·À«»  €Ì— ’ÕÌÕ° ÌÃ» √‰ Ì»œ√ »Õ—› ⁄—»Ì ›Ì «·”ÿ— —ﬁ„" & line_no
Else
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
End If
End Sub
Sub Constant_Type()
If Get_token() = "—ﬁ„Ì" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "—ﬁ„Ì" & vbNewLine
  If Get_token() = "’ÕÌÕ" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "’ÕÌÕ" & vbNewLine
    If Get_token() = "=" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "=" & vbNewLine
      If Not InStr(Constants & Variables, token(i - 5)) Then   'ÌÃ» «·›Õ’ ›Ì Ã„Ì⁄ «”„«¡ «·„ €Ì—«  Ê«·ÀÊ«»  Õ Ï ·« ÌﬂÊ‰ Â‰«ﬂ ·»”
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<—ﬁ„ ’ÕÌÕ>" & vbNewLine
        Scope = Scope + 1
        Integer_Number
        Scope = Scope - 1
        Integer_Constants = Integer_Constants & " " & token(i - 6) & " = " & token(i - 1)
        Constants = Constants & " " & token(i - 6)
      Else
        MsgBox "«”„ «·À«»  «·–Ì  Õ«Ê·  ⁄—Ì›Â ﬁœ ⁄—› ›Ì„« ”»ﬁ ·« Ì„ﬂ‰  ⁄—Ì› ‰›” «·«”„ „— Ì‰"
      End If
    Else
       MsgBox "ÌÃ» √‰ ÌﬂÊ‰ Â‰«ﬂ ﬁÌ„… ≈”‰«œ ··À«» "
    End If
  ElseIf token(i - 1) = "⁄‘—Ì" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "⁄‘—Ì" & vbNewLine
    If Get_token() = "=" Then
       Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "=" & vbNewLine
       Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<—ﬁ„ ⁄‘—Ì>" & vbNewLine
       Scope = Scope + 1
       Fractional_Number
       Scope = Scope - 1
       Fractional_Constants = Fractional_Constants & " " & token(i - 8) & " = " & token(i - 1) & token(i - 2) & token(i - 3)
       Constants = Constants & " " & token(i - 8)
    Else
       MsgBox "ÌÃ» √‰ ÌﬂÊ‰ Â‰«ﬂ ﬁÌ„… ≈”‰«œ ··À«» "
    End If
  End If
ElseIf token(i - 1) = "‰’¯Ì" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰’¯Ì" & vbNewLine
  If Get_token() = "=" Then
       Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "=" & vbNewLine
       Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‰’>" & vbNewLine
       Scope = Scope + 1
       Strings
       Scope = Scope - 1
       String_Constants = String_Constants & " " & token(i - 7) & " = " & token(i - 1) & token(i - 2) & token(i - 3)
       Constants = Constants & " " & token(i - 7)
  Else
       MsgBox "ÌÃ» √‰ ÌﬂÊ‰ Â‰«ﬂ ﬁÌ„… ≈”‰«œ ··À«» "
  End If
Else
  MsgBox "ÌÃ» √‰  Õœœ ‰Ê⁄ «·À«» "
End If
End Sub
Sub Integer_Number()
Dim j As Integer
Dim n As String
If Len(Get_token()) > 7 Then
  MsgBox "ÿÊ· «·—ﬁ„ «·’ÕÌÕ ÌÃ» √‰ ·« Ì Ã«Ê“ 7 Œ«‰« "
Else
  For j = 1 To Len(token(i - 1))
    n = Mid(token(i - 1), j, 1)
    If Not (n = "0" Or n = "1" Or n = "2" Or n = "3" Or n = "4" Or n = "5" Or n = "6" Or n = "7" Or n = "8" Or n = "9") Then
      MsgBox "«·—ﬁ„ «·’ÕÌÕ ÌÃ» √‰ ÌÕ ÊÌ ›ﬁÿ ⁄·Ï √—ﬁ«„ „‰ 0 - 9"
    End If
  Next j
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
End If
End Sub
Sub Fractional_Number()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<—ﬁ„ ’ÕÌÕ>" & vbNewLine
Scope = Scope + 1
Integer_Number
Scope = Scope - 1
If Get_token() = "." Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "." & vbNewLine
  Scope = Scope + 1
  Integer_Number
  Scope = Scope - 1
Else
  MsgBox "ÌÃ» √‰ ÌÕ ÊÌ «·—ﬁ„ «·⁄‘—Ì ⁄·Ï ›«’·… ⁄‘—Ì…"
End If
End Sub
Sub Strings()
If Get_token() = Chr(34) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & Chr(34) & vbNewLine
  If Get_token() = Chr(9) Then
    MsgBox "·« Ì„ﬂ‰ √‰  »œ√ ”ÿ— ÃœÌœ œÊ‰ √‰  ‰ÂÌ «·‰’"
  Else
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    If Not Get_token() = Chr(34) Then
      MsgBox "ÌÃ» √‰  ‰ÂÌ «·‰’ »⁄·«„… «· ‰’Ì’" & Chr(34)
    Else
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & Chr(34) & vbNewLine
    End If
  End If
Else
  MsgBox "ÌÃ» √‰ Ì»œ√ «·‰’ »⁄·«„… «· ‰’Ì’" & Chr(34)
End If
End Sub
Sub Variable_Deleration()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "⁄—¯ˆ›" & vbNewLine
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<√”„«¡ „ €Ì—« >" & vbNewLine
Scope = Scope + 1
var_start = i
Variable_Names
Scope = Scope - 1
If token(i - 1) = "ﬂ„ €Ì—" Or token(i - 1) = "ﬂ„ €Ì—« " Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‰Ê⁄ «·„ €Ì—>" & vbNewLine
  Scope = Scope + 1
  Variable_Type
  Scope = Scope - 1
End If

End Sub
Sub Variable_Names()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ „ €Ì—>" & vbNewLine
Scope = Scope + 1
Veriable_Name
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<„ €Ì—«  √Œ—Ï>" & vbNewLine
Scope = Scope + 1
Other_Variables
Scope = Scope - 1
End Sub
Sub Veriable_Name()
If Not Get_token() Like "[¡-Ì]*" Then
   MsgBox "«”„ «·„ €Ì— €Ì— ’ÕÌÕ° ÌÃ» √‰ Ì»œ√ »Õ—› ⁄—»Ì ›Ì «·”ÿ— —ﬁ„" & line_no
Else
   Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
End If
End Sub

Sub Other_Variables()
If Get_token() = "ﬂ„ €Ì—" Or token(i - 1) = "ﬂ„ €Ì—« " Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
  Exit Sub
ElseIf token(i - 1) = "°" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "°" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ „ €Ì—>" & vbNewLine
  Scope = Scope + 1
  Veriable_Name
  Scope = Scope - 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<„ €Ì—«  √Œ—Ï>" & vbNewLine
  Scope = Scope + 1
  Other_Variables
  Scope = Scope - 1
End If
End Sub
Sub Variable_Type()
Dim j As Integer
If Get_token() = "—ﬁ„Ì" Or token(i - 1) = "—ﬁ„Ì…" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  If Get_token() = "’ÕÌÕ" Or token(i - 1) = "’ÕÌÕ…" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    For j = var_start To i - 4 Step 2
      Integer_Variables = Integer_Variables & " " & token(j) & " = 0"
      Variables = Variables & " " & token(j)
    Next j
  ElseIf token(i - 1) = "⁄‘—Ì" Or token(i - 1) = "⁄‘—Ì…" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    For j = var_start To i - 4 Step 2
      Fractional_Variables = Fractional_Variables & " " & token(j) & " = 0"
      Variables = Variables & " " & token(j)
    Next j
  End If
ElseIf token(i - 1) = "‰’¯Ì" Or token(i - 1) = "‰’¯Ì…" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  For j = var_start To i - 3 Step 2
    String_Variables = String_Variables & " " & token(j) & " = " & Chr(34) & Chr(34)
    Variables = Variables & " " & token(j)
  Next j
End If
End Sub

Sub if_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "≈–«" & vbNewLine
If Get_token() = "ﬂ«‰" Or token(i - 1) = "ﬂ«‰ " Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  If Get_token() = "(" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‘—ÿ>" & vbNewLine
    Scope = Scope + 1
    Condition
    Scope = Scope - 1
    If Get_token() = ")" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
      If Get_token() = "‰›–" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›–" & vbNewLine
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
        Scope = Scope + 1
        Statment_List
        Scope = Scope - 1
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… ≈–«>" & vbNewLine
        Scope = Scope + 1
        if_Rest
        Scope = Scope - 1
      End If
    End If
  End If
End If
End Sub
Sub if_Rest()
If token(i - 1) = "‰Â«Ì…" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
  If Get_token() = "≈–«" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "≈–«" & vbNewLine
    Exit Sub
  End If
ElseIf token(i - 1) = "Ê" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "Ê" & vbNewLine
  If Get_token() = "€Ì—" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "€Ì—" & vbNewLine
    If Get_token() = "–·ﬂ" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "–·ﬂ" & vbNewLine
      If Get_token() = "‰›–" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›–" & vbNewLine
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
        Scope = Scope + 1
        Statment_List
        Scope = Scope - 1
        If token(i - 1) = "‰Â«Ì…" Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
          If Get_token() = "≈–«" Then
            Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "≈–«" & vbNewLine
          End If
        End If
      End If
    End If
  End If
End If
 
End Sub

Sub while_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "»Ì‰„«" & vbNewLine
If Get_token() = "ÌﬂÊ‰" Or token(i - 1) = " ﬂÊ‰" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  If Get_token() = "(" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‘—ÿ>" & vbNewLine
    Scope = Scope + 1
    Condition
    Scope = Scope - 1
    If Get_token() = ")" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
      If Get_token() = "‰›–" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›–" & vbNewLine
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
        Scope = Scope + 1
        Statment_List
        Scope = Scope - 1
        If token(i - 1) = "‰Â«Ì…" Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
          If Get_token() = "»Ì‰„«" Then
            Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "»Ì‰„«" & vbNewLine
          End If
        End If
      End If
    End If
  End If
End If

End Sub
Sub do_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«›⁄·" & vbNewLine
If Get_token() = "«· «·Ì" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«· «·Ì" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
  Scope = Scope + 1
  Statment_List
  Scope = Scope - 1
  If token(i - 1) = "ÿ«·„«" Then 'ÌÃ» «÷«›… ÿ«·„« ›Ì rest of stmt ·ﬂÌ  ” ÿÌ⁄ «‰Â«¡ ﬁ«∆„… «·Ã„·
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "ÿ«·„«" & vbNewLine
    If Get_token() = "(" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‘—ÿ>" & vbNewLine
      Scope = Scope + 1
      Condition
      Scope = Scope - 1
      If Get_token() = ")" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
      End If
    End If
  End If
End If


End Sub

Sub for_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "„‰" & vbNewLine
If Get_token = "ﬁÌ„…" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "ﬁÌ„…" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<Ã„·… ≈”‰«œ>" & vbNewLine
  Scope = Scope + 1
  Equality_stamt
  Scope = Scope - 1
  If Get_token() = "≈·Ï" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "≈·Ï" & vbNewLine
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
    Scope = Scope + 1
    Expression
    Scope = Scope - 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "»„ﬁœ«—" & vbNewLine
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<“Ì«œ…\‰ﬁ’«‰>" & vbNewLine
    Scope = Scope + 1
    Dec_Inc
    Scope = Scope - 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
    Scope = Scope + 1
    Expression
    Scope = Scope - 1
    If Get_token() = "‰›–" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›–" & vbNewLine
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
      Scope = Scope + 1
      Statment_List
      Scope = Scope - 1
      If token(i - 1) = "‰Â«Ì…" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
        If Get_token() = "„‰" Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "„‰" & vbNewLine
        End If
      End If
    End If
  End If
End If

End Sub
Sub Dec_Inc()
If Get_token() = "»„ﬁœ«—" Then
  If Get_token() = "“Ì«œ…" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "“Ì«œ…" & vbNewLine
  ElseIf token(i - 1) = "‰ﬁ’«‰" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰ﬁ’«‰" & vbNewLine
  End If
End If
End Sub
Sub switch_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "÷⁄" & vbNewLine
If Get_token() = "(" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ „ €Ì—>" & vbNewLine
  Scope = Scope + 1
  If InStr(Variables, Get_token()) Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    Scope = Scope - 1
    If Get_token() = ")" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
      If Get_token() = "··«Œ Ì«—" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "··«Œ Ì«—" & vbNewLine
        If Get_token() = Chr(9) Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Õ«·« >" & vbNewLine
          Scope = Scope + 1
          Cases_List
          Scope = Scope - 1
          While token(i) = Chr(9)
            Get_token
          Wend
          If Get_token() = "‰Â«Ì…" Then
            Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
            If Get_token() = "«·«Œ Ì«—" Then
              Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«·«Œ Ì«—" & vbNewLine
            End If
          End If
        End If
      End If
    End If
  End If
End If
  

End Sub
Sub Cases_List()
If Get_token() = "‰Â«Ì…" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
  Exit Sub
ElseIf token(i - 1) = "›Ì" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "›Ì" & vbNewLine
  If Get_token() = "Õ«·…" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "Õ«·…" & vbNewLine
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
    Scope = Scope + 1
    Expression
    Scope = Scope - 1
    If Get_token() = ":" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ":" & vbNewLine
      If Get_token() = "‰›–" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›–" & vbNewLine
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
        Scope = Scope + 1
        Statment_List
        Scope = Scope - 1
        If token(i - 1) = "‰Â«Ì…" Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
        End If
      End If
    End If
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Õ«·« >" & vbNewLine
  Scope = Scope + 1
  Cases_List
  Scope = Scope - 1
ElseIf token(i - 1) = "«·Õ«·…" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«·Õ«·…" & vbNewLine
  If Get_token() = "«·«› —«÷Ì…" Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«·«› —«÷Ì…" & vbNewLine
    If Get_token() = ":" Then
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ":" & vbNewLine
      If Get_token() = "‰›–" Then
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›–" & vbNewLine
        Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<ﬁ«∆„… «·Ã„·>" & vbNewLine
        Scope = Scope + 1
        Statment_List
        Scope = Scope - 1
        If token(i - 1) = "‰Â«Ì…" Then
          Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰Â«Ì…" & vbNewLine
        End If
      End If
    End If
  End If
ElseIf token(i - 1) = Chr(9) Then
  Cases_List
End If

End Sub

Sub go_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«–Â»" & vbNewLine
If Get_token() = "≈·Ï" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "≈·Ï" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ ⁄‰Ê«‰>" & vbNewLine
  Scope = Scope + 1
' check Label_Name
  If Not InStr(Labels, Get_token()) Then
    MsgBox "·« Ì„ﬂ‰ «·–Â«» ≈·Ï ⁄‰Ê«‰ €Ì— „ÊÃÊœ"
  Else
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    Scope = Scope - 1
  End If
End If

End Sub

Sub Label_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ ⁄‰Ê«‰>" & vbNewLine
Scope = Scope + 1
'add token(i-1) to labels tabel
Labels = Labels & " " & token(i - 1) & " = " & line_no
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
Scope = Scope - 1
i = i + 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ":" & vbNewLine
End Sub
Sub Equality_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ „ €Ì—>" & vbNewLine
Scope = Scope + 1
If InStr(Variables, Get_token()) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Scope = Scope - 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<⁄„·Ì… ≈”‰«œ>" & vbNewLine
  Scope = Scope + 1
  Assignment_Operator
  Scope = Scope - 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
  Scope = Scope + 1
  Expression
  Scope = Scope - 1
Else
  MsgBox "·« Ì„ﬂ‰ «·«”‰«œ ·„ €Ì— €Ì— „⁄—¯›"
End If
End Sub
Sub Assignment_Operator()
If Not (Get_token() = "=" Or token(i - 1) = "+=" Or token(i - 1) = "-=" Or token(i - 1) = "*=" Or token(i - 1) = "\=" Or token(i - 1) = "^=" Or token(i - 1) = "%=") Then
  MsgBox "⁄„·Ì… «·«”‰«œ Â–Â €Ì— ’ÕÌÕ…° «·⁄„·Ì«  «·„ Ê›—… ÂÌ ›ﬁÿ = += -= *= \= ^= %= ›Ì «·”ÿ— —ﬁ„" & line_no
Else
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
End If
End Sub
Sub Expression()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<≈‘«—…>" & vbNewLine
Scope = Scope + 1
Sign
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<⁄»«—…>" & vbNewLine
Scope = Scope + 1
Phrase
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… «· ⁄»Ì—>" & vbNewLine
Scope = Scope + 1
Expression_Rest
Scope = Scope - 1
End Sub
Sub Sign()
Dim ln As Integer
ln = line_no
If Not (Get_token() = "+" Or token(i - 1) = "-") Then
  i = i - 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
  If line_no > ln Then
    line_no = line_no - 1
  End If
Else
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
End If
End Sub
Sub Phrase()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<„⁄«„· —Ì«÷Ì>" & vbNewLine
Scope = Scope + 1
Arith_Factor
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… «·⁄»«—…>" & vbNewLine
Scope = Scope + 1
Phrase_Rest
Scope = Scope - 1
End Sub
Sub Expression_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = "+" Or token(i - 1) = "-" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
  Scope = Scope + 1
  Expression
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If
End Sub

Sub Arith_Factor()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<„⁄«„·>" & vbNewLine
Scope = Scope + 1
Factor
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… «·„⁄«„·>" & vbNewLine
Scope = Scope + 1
Factor_Rest
Scope = Scope - 1
End Sub
Sub Phrase_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = "*" Or token(i - 1) = "\" Or token(i - 1) = "%" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<⁄»«—…>" & vbNewLine
  Scope = Scope + 1
  Phrase
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If
End Sub
Sub Factor_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = "^" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<„⁄«„· —Ì«÷Ì>" & vbNewLine
  Scope = Scope + 1
  Arith_Factor
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If
End Sub
Sub Factor()
Dim ln As Integer
ln = line_no
Dim n As String
Dim j As Integer
n = Mid(Get_token(), 1, 1)
If token(i - 1) = "(" Then
  paranthes_counter = paranthes_counter + 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "(" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
  Scope = Scope + 1
  Expression
  Scope = Scope - 1
  If Get_token() = ")" Then
    paranthes_counter = paranthes_counter - 1
  Else
    MsgBox "·ﬁœ › Õ  ﬁÊ” œÊ‰ √‰  €·ﬁÂ"
  End If
ElseIf token(i - 1) = ")" Then
  If paranthes_counter = 0 Then
    MsgBox "ÌÃ» √‰ ·«  €·ﬁ ﬁÊ” »œÊ‰ √‰  › ÕÂ"
  Else
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ")" & vbNewLine
    paranthes_counter = paranthes_counter - 1
  End If
ElseIf token(i - 1) = "‰›Ì" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "‰›Ì" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
  Scope = Scope + 1
  Expression
  Scope = Scope - 1
ElseIf InStr(Variables, token(i - 1)) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ „ €Ì—>" & vbNewLine
  Scope = Scope + 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Scope = Scope - 1
ElseIf InStr(Constants, token(i - 1)) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ À«» >" & vbNewLine
  Scope = Scope + 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Scope = Scope - 1
ElseIf token(i - 1) = "’Õ" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "’Õ" & vbNewLine
ElseIf token(i - 1) = "Œÿ√" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "Œÿ√" & vbNewLine
ElseIf n Like "[0-9]" Then
  If Len(token(i - 1)) > 7 Then
    MsgBox "ÿÊ· «·—ﬁ„ «·’ÕÌÕ ÌÃ» √‰ ·« Ì Ã«Ê“ 7 Œ«‰« "
  Else
    For j = 1 To Len(token(i - 1))
      n = Mid(token(i - 1), j, 1)
      If Not (n = "0" Or n = "1" Or n = "2" Or n = "3" Or n = "4" Or n = "5" Or n = "6" Or n = "7" Or n = "8" Or n = "9") Then
        MsgBox "«·—ﬁ„ «·’ÕÌÕ ÌÃ» √‰ ÌÕ ÊÌ ›ﬁÿ ⁄·Ï √—ﬁ«„ „‰ 0 - 9 ÊÌÃ» √‰ ·« Ì ’· »«”„ ¬Œ— ≈·« »⁄„·Ì… Õ”«»Ì…"
      End If
    Next j
  End If
  If Get_token() = "." Then
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<—ﬁ„ ⁄‘—Ì>" & vbNewLine
    Scope = Scope + 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 2) & vbNewLine
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "." & vbNewLine
    If Len(Get_token()) > 7 Then
      MsgBox "ÿÊ· «·Ã“¡ «·⁄‘—Ì „‰ «·—ﬁ„ ÌÃ» √‰ ·« Ì Ã«Ê“ 7 Œ«‰« "
    Else
      For j = 1 To Len(token(i - 1))
        n = Mid(token(i - 1), j, 1)
        If Not (n = "0" Or n = "1" Or n = "2" Or n = "3" Or n = "4" Or n = "5" Or n = "6" Or n = "7" Or n = "8" Or n = "9") Then
          MsgBox "«·—ﬁ„ «·’ÕÌÕ ÌÃ» √‰ ÌÕ ÊÌ ›ﬁÿ ⁄·Ï √—ﬁ«„ „‰ 0 - 9 ÊÌÃ» √‰ ·« Ì ’· »«”„ ¬Œ— ≈·« »⁄„·Ì… Õ”«»Ì…"
        End If
      Next j
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
      Scope = Scope - 1
    End If
  Else
    i = i - 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<—ﬁ„ ’ÕÌÕ>" & vbNewLine
    Scope = Scope + 1
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    Scope = Scope - 1
    If line_no > ln Then
      line_no = line_no - 1
    End If
  End If
ElseIf token(i - 1) = Chr(34) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‰’>" & vbNewLine
  Scope = Scope + 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & Chr(34) & vbNewLine
  If Get_token() = Chr(9) Then
    MsgBox "·« Ì„ﬂ‰ √‰  »œ√ ”ÿ— ÃœÌœ œÊ‰ √‰  ‰ÂÌ «·‰’"
  Else
    Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
    If Not Get_token() = Chr(34) Then
      MsgBox "ÌÃ» √‰  ‰ÂÌ «·‰’ »⁄·«„… «· ‰’Ì’" & Chr(34)
    Else
      Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & Chr(34) & vbNewLine
      Scope = Scope - 1
    End If
  End If
ElseIf InStr("+ - * \ ^ %", token(i - 1)) Then
  MsgBox "·« Ì„ﬂ‰ √‰  ÷⁄ ⁄„·Ì Ì‰ Õ”«»Ì Ì‰ »Ã«‰» »⁄÷"
Else
  MsgBox "«”„ „ €Ì— «Ê À«»  €Ì— „⁄—› œ«Œ· ⁄„·Ì… «·«”‰«œ"
End If
End Sub

Sub Condition()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‘—ÿ „»«‘—>" & vbNewLine
Scope = Scope + 1
Direct_Condition
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… «·‘—ÿ>" & vbNewLine
Scope = Scope + 1
Condition_Rest
Scope = Scope - 1
End Sub
Sub Direct_Condition()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
Scope = Scope + 1
Expression
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… ‘—ÿ „»«‘—>" & vbNewLine
Scope = Scope + 1
Direct_Condition_Rest
Scope = Scope - 1
End Sub
Sub Condition_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = "Ê" Or token(i - 1) = "√Ê" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<‘—ÿ>" & vbNewLine
  Scope = Scope + 1
  Condition
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If
End Sub
Sub Direct_Condition_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = ">" Or token(i - 1) = "<" Or token(i - 1) = ">=" Or token(i - 1) = "<=" Or token(i - 1) = "==" Or token(i - 1) = "!=" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
  Scope = Scope + 1
  Expression
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If
End Sub
Sub Print_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "«ÿ»⁄" & vbNewLine
If Get_token() = "<<" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<<" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<”·”·… «ÿ»⁄>" & vbNewLine
  Scope = Scope + 1
  Print_List
  Scope = Scope - 1
Else
  MsgBox "ÌÃ» √‰  ›’· „⁄ÿÌ«  «·ÿ»«⁄… »«·«‘«—… <<"
End If
End Sub
Sub Print_List()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "< ⁄»Ì—>" & vbNewLine
Scope = Scope + 1
Expression
Scope = Scope - 1
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… ”·”·… «ÿ»⁄>" & vbNewLine
Scope = Scope + 1
Print_List_Rest
Scope = Scope - 1
End Sub
Sub Print_List_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = "<<" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<<" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<”·”·… «ÿ»⁄>" & vbNewLine
  Scope = Scope + 1
  Print_List
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If

End Sub
Sub Input_stamt()
Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "√œŒ·" & vbNewLine
If Get_token() = ">>" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ">>" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<”·”·… √œŒ·>" & vbNewLine
  Scope = Scope + 1
  Input_List
  Scope = Scope - 1
Else
  MsgBox "ÌÃ» √‰  ›’· „⁄ÿÌ«  «·ÿ»«⁄… »«·«‘«—… >>"
End If
End Sub
Sub Input_List()
If InStr(Variables, Get_token()) Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<«”„ „ €Ì—>" & vbNewLine
  Scope = Scope + 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & token(i - 1) & vbNewLine
  Scope = Scope - 1
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<»ﬁÌ… ”·”·… √œŒ·>" & vbNewLine
  Scope = Scope + 1
  Input_List_Rest
  Scope = Scope - 1
Else
  MsgBox "·« Ì„ﬂ‰ √‰  œŒ· ”ÊÏ ·„ €Ì—«  „⁄—¯›…"
End If
End Sub
Sub Input_List_Rest()
Dim ln As Integer
ln = line_no
If Get_token() = ">>" Then
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & ">>" & vbNewLine
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "<”·”·… √œŒ·>" & vbNewLine
  Scope = Scope + 1
  Input_List
  Scope = Scope - 1
Else
  i = i - 1
  If line_no > ln Then
    line_no = line_no - 1
  End If
  Parsed_Tree = Parsed_Tree & Add_Tabs(Scope, "") & "·« ‘Ì¡" & vbNewLine
End If
End Sub


