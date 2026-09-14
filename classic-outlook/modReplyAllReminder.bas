Attribute VB_Name = "modReplyAllReminder"
Option Explicit

' =====================================================================
'  Reply-All Reminder - Classic Outlook for Windows
' ---------------------------------------------------------------------
'  Warns when you Reply to a message that had other recipients, listing
'  exactly who is about to be dropped from the thread.
'
'  Install: see README.md. The Application_ItemSend stub in
'  ThisOutlookSession calls CheckReplyAll below.
' =====================================================================

' --------------------------- CONFIG ----------------------------------

' Extra addresses that count as "you" (aliases, shared mailboxes you own).
' Your Outlook account addresses are detected automatically.
Private Const MY_EXTRA_ADDRESSES As String = ""          ' e.g. "me@old.com;me@alias.com"

' Never warn about these addresses or domain fragments.
Private Const IGNORE_FRAGMENTS As String = "noreply@;no-reply@;donotreply@"

' Warn only when at least this many people would be left out.
Private Const MIN_DROPPED As Long = 1

' Header labels in the quoted original. Add your locale's if needed.
Private Const FROM_LABELS As String = "from;von;de;da;van;fran;fra;od"
Private Const SENT_LABELS As String = "sent;date;gesendet;envoye;enviado;inviato;verzonden;skickat;sendt;wyslano"
Private Const SUBJECT_LABELS As String = "subject;betreff;objet;asunto;oggetto;onderwerp;amne;emne;aihe;temat"
Private Const TO_LABELS As String = "to;an;a;para;aan;til;till;do;komu"
Private Const CC_LABELS As String = "cc;copy;kopie;copia;kopia"

Private Const PR_IN_REPLY_TO_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
Private Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

' ---------------------------------------------------------------------
'  Entry point - called from Application_ItemSend
' ---------------------------------------------------------------------
Public Sub CheckReplyAll(ByVal Item As Object, ByRef Cancel As Boolean)
    On Error GoTo Fallthrough      ' never block a send because of a bug here

    If Item.Class <> olMail Then Exit Sub
    If Not IsReply(Item) Then Exit Sub

    Dim block As Collection
    Set block = ExtractHeaderBlock(Item.Body)
    If block Is Nothing Then Exit Sub

    Dim original As Object          ' Dictionary: key -> display text
    Set original = ParseParticipants(block)
    If original.Count = 0 Then Exit Sub

    Dim current As Object
    Set current = CurrentRecipientKeys(Item)

    Dim mine As Object
    Set mine = MyAddressKeys()

    Dim dropped As String, droppedCount As Long
    Dim k As Variant
    For Each k In original.Keys
        If Not current.Exists(k) And Not mine.Exists(k) And Not IsIgnored(CStr(k)) Then
            droppedCount = droppedCount + 1
            dropped = dropped & "    " & original(k) & vbCrLf
        End If
    Next k

    If droppedCount < MIN_DROPPED Then Exit Sub

    Dim msg As String
    msg = "You clicked Reply, not Reply All." & vbCrLf & vbCrLf & _
          "These people were on the original message and will NOT " & _
          "receive this response:" & vbCrLf & vbCrLf & dropped & vbCrLf & _
          "Send anyway?"

    If MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, _
              "Reply-All Reminder") = vbNo Then
        Cancel = True
        Item.Display
    End If

Fallthrough:
End Sub

' ---------------------------------------------------------------------
'  Is this a reply? In-Reply-To is set on replies but not on forwards,
'  which makes this locale-proof (unlike checking for an "RE:" prefix).
' ---------------------------------------------------------------------
Private Function IsReply(ByVal Item As Object) As Boolean
    On Error Resume Next
    Dim v As String
    v = Item.PropertyAccessor.GetProperty(PR_IN_REPLY_TO_ID)
    IsReply = (Err.Number = 0) And (Len(Trim$(v)) > 0)
    Err.Clear
End Function

' ---------------------------------------------------------------------
'  The From/Sent/To/Cc/Subject block above the most recent quoted message
' ---------------------------------------------------------------------
Private Function ExtractHeaderBlock(ByVal body As String) As Collection
    Dim lines() As String
    lines = Split(Replace(body, vbCrLf, vbLf), vbLf)

    Dim i As Long, j As Long
    Dim block As Collection

    For i = LBound(lines) To UBound(lines)
        If MatchesLabel(lines(i), FROM_LABELS) Then
            Set block = New Collection
            For j = i To MinL(i + 9, UBound(lines))
                If Len(Trim$(lines(j))) = 0 And block.Count >= 2 Then Exit For
                block.Add lines(j)
                If MatchesLabel(lines(j), SUBJECT_LABELS) Then Exit For
            Next j
            If block.Count >= 2 Then Set ExtractHeaderBlock = block
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------------
'  Participants of the original, excluding its sender (already your To:)
' ---------------------------------------------------------------------
Private Function ParseParticipants(ByVal block As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1     ' TextCompare

    Dim idx As Long, line As Variant
    For Each line In block
        idx = idx + 1
        If idx > 1 Then
            If Not MatchesLabel(CStr(line), FROM_LABELS) _
               And Not MatchesLabel(CStr(line), SENT_LABELS) _
               And Not MatchesLabel(CStr(line), SUBJECT_LABELS) Then

                Dim labelled As Boolean
                labelled = MatchesLabel(CStr(line), TO_LABELS) Or _
                           MatchesLabel(CStr(line), CC_LABELS)

                ' An unlabelled line counts only if it really reads like a
                ' recipient list - this keeps unlisted locales working.
                If labelled Or (Not LooksLikeTimestamp(CStr(line)) _
                                And LooksLikeRecipientList(StripLabel(CStr(line)))) Then
                    AddEntries result, StripLabel(CStr(line))
                End If
            End If
        End If
    Next line

    Set ParseParticipants = result
End Function

Private Sub AddEntries(ByRef dict As Object, ByVal value As String)
    Dim parts() As String, p As Variant, entry As String, key As String
    parts = Split(value, ";")

    For Each p In parts
        entry = Trim$(Replace(Replace(CStr(p), Chr$(34), ""), "'", ""))
        If Len(entry) > 0 Then
            key = EntryKey(entry)
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then dict.Add key, entry
            End If
        End If
    Next p
End Sub

' Email address if there is one, otherwise the display name.
Private Function EntryKey(ByVal entry As String) As String
    Dim lt As Long, gt As Long, inner As String
    lt = InStr(entry, "<")
    gt = InStr(entry, ">")

    If lt > 0 And gt > lt Then
        inner = Trim$(Mid$(entry, lt + 1, gt - lt - 1))
        If InStr(inner, "@") > 0 Then
            EntryKey = LCase$(inner)
            Exit Function
        End If
    End If

    If InStr(entry, "@") > 0 Then
        EntryKey = LCase$(Trim$(entry))
    ElseIf Len(entry) >= 2 And Len(entry) <= 80 Then
        EntryKey = LCase$(Trim$(entry))
    End If
End Function

' ---------------------------------------------------------------------
'  Who is on the reply right now
' ---------------------------------------------------------------------
Private Function CurrentRecipientKeys(ByVal Item As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim r As Object, smtp As String
    For Each r In Item.Recipients
        smtp = RecipientSmtp(r)
        If Len(smtp) > 0 Then
            If Not d.Exists(LCase$(smtp)) Then d.Add LCase$(smtp), 1
        End If
        If Len(r.Name) > 0 Then
            If Not d.Exists(LCase$(r.Name)) Then d.Add LCase$(r.Name), 1
        End If
    Next r

    Set CurrentRecipientKeys = d
End Function

Private Function RecipientSmtp(ByVal r As Object) As String
    On Error Resume Next
    RecipientSmtp = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    If Err.Number <> 0 Or Len(RecipientSmtp) = 0 Then
        Err.Clear
        RecipientSmtp = r.AddressEntry.GetExchangeUser.PrimarySmtpAddress
    End If
    If Err.Number <> 0 Then
        Err.Clear
        RecipientSmtp = r.Address
    End If
    If InStr(RecipientSmtp, "@") = 0 Then RecipientSmtp = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------
'  Your own addresses - every account, plus MY_EXTRA_ADDRESSES
' ---------------------------------------------------------------------
Private Function MyAddressKeys() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    On Error Resume Next

    Dim acct As Object
    For Each acct In Application.Session.Accounts
        If Len(acct.SmtpAddress) > 0 Then
            If Not d.Exists(LCase$(acct.SmtpAddress)) Then d.Add LCase$(acct.SmtpAddress), 1
        End If
    Next acct

    Dim me_ As String
    me_ = Application.Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
    If Len(me_) > 0 Then If Not d.Exists(LCase$(me_)) Then d.Add LCase$(me_), 1

    Dim myName As String
    myName = Application.Session.CurrentUser.Name
    If Len(myName) > 0 Then If Not d.Exists(LCase$(myName)) Then d.Add LCase$(myName), 1

    Err.Clear

    Dim extra() As String, e As Variant
    extra = Split(MY_EXTRA_ADDRESSES, ";")
    For Each e In extra
        If Len(Trim$(CStr(e))) > 0 Then
            If Not d.Exists(LCase$(Trim$(CStr(e)))) Then d.Add LCase$(Trim$(CStr(e))), 1
        End If
    Next e

    Set MyAddressKeys = d
End Function

' ---------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------
Private Function IsIgnored(ByVal key As String) As Boolean
    Dim frags() As String, f As Variant
    frags = Split(IGNORE_FRAGMENTS, ";")
    For Each f In frags
        If Len(Trim$(CStr(f))) > 0 Then
            If InStr(1, key, Trim$(CStr(f)), vbTextCompare) > 0 Then
                IsIgnored = True
                Exit Function
            End If
        End If
    Next f
End Function

Private Function LabelOf(ByVal line As String) As String
    Dim c As Long
    c = InStr(line, ":")
    If c > 1 And c <= 26 Then LabelOf = LCase$(Trim$(Left$(line, c - 1)))
End Function

Private Function MatchesLabel(ByVal line As String, ByVal labelList As String) As Boolean
    Dim label As String
    label = LabelOf(line)
    If Len(label) = 0 Then Exit Function
    MatchesLabel = (InStr(1, ";" & labelList & ";", ";" & label & ";", vbTextCompare) > 0)
End Function

Private Function StripLabel(ByVal line As String) As String
    Dim c As Long
    c = InStr(line, ":")
    If c > 0 And Len(LabelOf(line)) > 0 Then
        StripLabel = Trim$(Mid$(line, c + 1))
    Else
        StripLabel = Trim$(line)
    End If
End Function

' A 4-digit year or a clock time means this is a date line, not recipients.
Private Function LooksLikeTimestamp(ByVal line As String) As Boolean
    Dim i As Long, run As Long

    For i = 1 To Len(line)
        If IsDigitChar(Mid$(line, i, 1)) Then
            run = run + 1
            If run >= 4 Then
                LooksLikeTimestamp = True
                Exit Function
            End If
        Else
            run = 0
        End If
    Next i

    For i = 2 To Len(line) - 1
        If Mid$(line, i, 1) = ":" Then
            If IsDigitChar(Mid$(line, i - 1, 1)) And IsDigitChar(Mid$(line, i + 1, 1)) Then
                LooksLikeTimestamp = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsDigitChar(ByVal c As String) As Boolean
    IsDigitChar = (c >= "0" And c <= "9")
End Function

Private Function LooksLikeRecipientList(ByVal value As String) As Boolean
    Dim parts() As String, p As Variant, plausible As Long, t As String
    parts = Split(value, ";")
    If UBound(parts) < 1 Then Exit Function

    For Each p In parts
        t = Trim$(CStr(p))
        If Len(t) > 0 Then
            If InStr(t, "@") > 0 Then
                plausible = plausible + 1
            ElseIf Len(t) <= 60 And Not LooksLikeTimestamp(t) Then
                plausible = plausible + 1
            End If
        End If
    Next p

    LooksLikeRecipientList = (plausible >= 2)
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function
