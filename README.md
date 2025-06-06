# CodeAccess

```vba
ФОРМА АВТОРИЗАЦИЯ
Private Sub Form_Load()
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim Текущая_дата As Date
    Dim Логин As String
    Dim Дата_последней_авторизации As Date
    
    Текущая_дата = Date
    
    strConn = "Provider=SQLOLEDB;Data Source=DO_ZAVRTA\MS;Initial Catalog=HotelSystem;Integrated Security=SSPI;"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn
    
    strSQL = "SELECT login, last_auth_date FROM [users] WHERE last_auth_date IS NOT NULL"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 1, 3
    
    Do While Not rs.EOF
        Логин = rs.Fields("login").Value
        Дата_последней_авторизации = rs.Fields("last_auth_date").Value
        
        If Текущая_дата - Дата_последней_авторизации > 30 Then
            Dim updateSQL As String
            updateSQL = "UPDATE [users] SET is_blocked = 1 WHERE login = '" & Replace(Логин, "'", "''") & "'"
            conn.Execute updateSQL
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    
    Set rs = Nothing
    Set conn = Nothing
End Sub
Private Sub Вход_Click()
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim Логин As String
    Dim Пароль As String
    Dim Права_доступа As String
    Dim Блокировка As Boolean
    Dim ПопыткиВхода As Integer
    Dim ПодтверждениеПароля As Boolean

    Логин = Nz(Me.Логин.Value, "")
    Пароль = Nz(Me.Пароль.Value, "")

    If Trim(Логин) = "" Or Trim(Пароль) = "" Then
        MsgBox "Введите логин и пароль"
        Exit Sub
    End If

    strConn = "Provider=SQLOLEDB;Data Source=DO_ZAVRTA\MS;Initial Catalog=HotelSystem;Integrated Security=SSPI;"

    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = strConn
    conn.Open

    strSQL = "SELECT * FROM [users] WHERE login = '" & Логин & "'"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 1, 3

    If rs.EOF Then
        MsgBox "Пользователь с таким логином не найден"
    Else
        Блокировка = Nz(rs("is_blocked"), False)
        If Блокировка Then
            MsgBox "Ваш аккаунт заблокирован. Обратитесь к администратору."
            rs.Close: conn.Close
            Exit Sub
        End If

        If rs("password") = Пароль Then
            ТекущийЛогин = Логин
            rs("login_attempts") = 0
            rs("last_auth_date") = Date
            ПодтверждениеПароля = Nz(rs("account_confirmed"), False)
            Права_доступа = rs("access_level")

            rs.Update

            MsgBox "Вы успешно авторизовались"
            If Not ПодтверждениеПароля Then
                MsgBox "Пожалуйста, смените пароль перед продолжением."
                DoCmd.OpenForm "Смена пароля"
            End If

            Select Case Права_доступа
                Case "Администратор"
                    DoCmd.OpenForm "Главное меню администратора"
                Case "Сотрудник"
                    DoCmd.OpenForm "Главное меню сотрудника"
                Case "Клиент"
                    DoCmd.OpenForm "Главное меню клиента"
                Case Else
                    MsgBox "Неизвестная роль пользователя"
            End Select

            DoCmd.Close acForm, Me.Name
        Else
            ПопыткиВхода = Nz(rs("login_attempts"), 0) + 1
            rs("login_attempts") = ПопыткиВхода

            If ПопыткиВхода >= 3 Then
                rs("is_blocked") = True
                MsgBox "Ваш аккаунт заблокирован из-за превышения количества неудачных попыток."
            Else
                MsgBox "Неверный пароль. Осталось попыток: " & (3 - ПопыткиВхода)
            End If

            rs.Update
        End If
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
Private Sub Выход_Click()
    Dim ответ As Integer
    ответ = MsgBox("Вы действительно хотите выйти?", vbYesNo)
    
    If ответ = vbYes Then
        DoCmd.Quit
    End If
End Sub
```

```vba
# ФОРМА ГЛАВНОЕ МЕНЮ (ЛЮБОЕ)
Private Sub Выход_Click()
    Dim ответ As Integer
    ответ = MsgBox("Вы действительно хотите выйти?", vbYesNo)
    
    If ответ = vbYes Then
        DoCmd.Quit
    End If
End Sub
Private Sub СменаПользователя_Click()
    Dim frm As AccessObject

    For Each frm In Application.CurrentProject.AllForms
        If frm.IsLoaded And frm.Name <> Me.Name Then
            DoCmd.Close acForm, frm.Name
        End If
    Next frm

    DoCmd.Close acForm, Me.Name

    DoCmd.OpenForm "Авторизация"
End Sub
```

```vba
# ФОРМА СМЕНА ПАРОЛЯ
Private Sub btnСохранить_Click()
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim СтарыйПароль As String
    Dim НовыйПароль As String
    Dim ПодтверждениеПароля As String
    
    СтарыйПароль = Nz(Me.СтарыйПароль.Value, "")
    НовыйПароль = Nz(Me.НовыйПароль.Value, "")
    ПодтверждениеПароля = Nz(Me.ПодтверждениеПароля.Value, "")

    If Trim(СтарыйПароль) = "" Or Trim(НовыйПароль) = "" Or Trim(ПодтверждениеПароля) = "" Then
        MsgBox "Заполните все поля."
        Exit Sub
    End If
    
    If НовыйПароль <> ПодтверждениеПароля Then
        MsgBox "Новый пароль и подтверждение не совпадают."
        Exit Sub
    End If

    strConn = "Provider=SQLOLEDB;Data Source=DO_ZAVRTA\MS;Initial Catalog=HotelSystem;Integrated Security=SSPI;"
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn

    strSQL = "SELECT * FROM [users] WHERE login = '" & Replace(ТекущийЛогин, "'", "''") & "'"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 1, 3

    If rs.EOF Then
        MsgBox "Пользователь не найден."
    ElseIf rs.Fields("password").Value <> СтарыйПароль Then
        MsgBox "Старый пароль введен неверно."
    Else
        rs.Fields("password").Value = НовыйПароль
        rs.Update
        MsgBox "Пароль успешно изменен."
        DoCmd.Close acForm, Me.Name
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
Private Sub Form_Close()
    Dim conn As Object
    Dim strConn As String
    Dim updateSQL As String

    If ТекущийЛогин = "" Then Exit Sub

    strConn = "Provider=SQLOLEDB;Data Source=DO_ZAVRTA\MS;Initial Catalog=HotelSystem;Integrated Security=SSPI;"
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn

    updateSQL = "UPDATE [users] SET account_confirmed = 1 WHERE login = '" & Replace(ТекущийЛогин, "'", "''") & "'"
    conn.Execute updateSQL

    conn.Close
    Set conn = Nothing
End Sub
```

```vba
# ФОРМА РЕДАКТИРОВАНИЕ/ДОБАВЛЕНИЯ ПОЛЬЗОВАТЕЛЯ
Private Sub login_BeforeUpdate(Cancel As Integer)
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim НовыйЛогин As String

    НовыйЛогин = Replace(Nz(Me.login.Value, ""), "'", "''")
    If Trim(НовыйЛогин) = "" Then Exit Sub

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB;Data Source=DO_ZAVRTA\MS;Initial Catalog=HotelSystem;Integrated Security=SSPI;"

    strSQL = "SELECT login FROM [users] WHERE login = '" & НовыйЛогин & "'"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn, 1, 1

    If Not rs.EOF Then
        MsgBox "Пользователь с таким логином уже существует. Введите другой логин.", vbExclamation
        Cancel = True
        Me.login.Undo
    End If

    rs.Close: conn.Close
    Set rs = Nothing: Set conn = Nothing
End Sub
Private Sub is_blocked_BeforeUpdate(Cancel As Integer)
    If Me.is_blocked = False Then
        Me.last_auth_date = Date
        Me.login_attempts = 0
    End If
End Sub
```

```vba
# ГЛОБАЛЬНЫЙ МОДУЛЬ
Public ТекущийЛогин As String

# САМОЕ ГЛАВНОЕ
    strConn = "Provider=SQLOLEDB;Data Source=ИМЯ_СЕРВЕРА;Initial Catalog=ИМЯ_БД;Integrated Security=SSPI;"
# Не  забывайте менять в этой строчке имя сервера и название вашей базы
```










