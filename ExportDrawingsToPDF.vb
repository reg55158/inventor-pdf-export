Imports System.IO
AddReference "System.Drawing"

Sub Main()
    Dim asmDoc As AssemblyDocument = ThisApplication.ActiveDocument
    Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
    Dim outFolder As String = System.IO.Path.Combine(asmFolder, "PDF Export")
    System.IO.Directory.CreateDirectory(outFolder)

    ' 1. Find drawings for parts AND assemblies named ###-#### ..., sorted by filename
    Dim assembliesFirst As Boolean = True   ' True = assembly drawings first, False = all mixed by number
    Dim sorted As New SortedDictionary(Of String, String)(StringComparer.Ordinal)
    Dim missingAsm As New List(Of String)
    Dim missingParts As New List(Of String)

    ' Top-level assembly plus everything it references
    Dim docsToCheck As New List(Of Document)
    docsToCheck.Add(asmDoc)
    For Each doc As Document In asmDoc.AllReferencedDocuments
        docsToCheck.Add(doc)
    Next

    For Each doc As Document In docsToCheck
        Dim isPart As Boolean = (doc.DocumentType = DocumentTypeEnum.kPartDocumentObject)
        Dim isAsm As Boolean = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
        If Not (isPart Or isAsm) Then Continue For

        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
        If Not (baseName Like "###-####*") Then Continue For

        Dim dwgPath As String = System.IO.Path.ChangeExtension(doc.FullFileName, ".idw")
        If System.IO.File.Exists(dwgPath) Then
            Dim sortKey As String = baseName
            If assembliesFirst Then sortKey = IIf(isAsm, "0", "1") & baseName
            If Not sorted.ContainsKey(sortKey) Then sorted.Add(sortKey, dwgPath)
        Else
            If isAsm Then missingAsm.Add(baseName) Else missingParts.Add(baseName)
        End If
    Next

    ' 2. Export each drawing to a PDF named after the drawing file
    Dim pdfAddIn As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
    Dim pdfFiles As New List(Of String)
    Dim skippedFiles As New List(Of String)

    For Each item In sorted
        Dim pdfPath As String = System.IO.Path.Combine(outFolder, System.IO.Path.GetFileNameWithoutExtension(item.Value) & ".pdf")

        ' Clear the old PDF first (before opening the drawing). Retry / Skip / Cancel if it's locked.
        Dim clr As Integer = ClearOldFile(pdfPath, True)
        If clr = 0 Then
            MessageBox.Show("Export cancelled. " & pdfFiles.Count & " drawings were exported before stopping.", "PDF export")
            Return
        End If
        If clr = 2 Then
            skippedFiles.Add(System.IO.Path.GetFileNameWithoutExtension(item.Value))
            Continue For
        End If

        ' Remember if the drawing was already open so we don't close it on you
        Dim wasOpen As Boolean = False
        For Each d As Document In ThisApplication.Documents
            If String.Equals(d.FullFileName, item.Value, StringComparison.OrdinalIgnoreCase) Then wasOpen = True
        Next

        Dim dwgDoc As Document = ThisApplication.Documents.Open(item.Value, False)

        Dim ctx As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext()
        ctx.Type = IOMechanismEnum.kFileBrowseIOMechanism
        Dim opts As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
        opts.Value("All_Color_AS_Black") = 0
        opts.Value("Vector_Resolution") = 4800
        opts.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
        Dim dm As DataMedium = ThisApplication.TransientObjects.CreateDataMedium()
        dm.FileName = pdfPath

        pdfAddIn.SaveCopyAs(dwgDoc, ctx, opts, dm)
        If Not wasOpen Then dwgDoc.Close(True)
        pdfFiles.Add(pdfPath)
    Next

    If pdfFiles.Count = 0 Then
        MessageBox.Show("No drawings were exported.", "PDF export")
        Return
    End If

    ' 3. Merge into one PDF using Acrobat (delete this block to skip merging)
    Dim outPath As String = System.IO.Path.Combine(asmFolder, System.IO.Path.GetFileNameWithoutExtension(asmDoc.FullFileName) & "_Drawings.pdf")

    If ClearOldFile(outPath, False) <> 1 Then
        MessageBox.Show("Merge cancelled. The individual PDFs are in:" & vbCrLf & outFolder, "PDF export")
        Return
    End If

    Dim stage As String = ""
    Dim merged As Object = Nothing
    Try
        stage = "creating Acrobat object"
        merged = CreateObject("AcroExch.PDDoc")

        stage = "opening first PDF"
        If Not CBool(merged.Open(pdfFiles(0))) Then Throw New Exception("Acrobat could not open " & pdfFiles(0))

        For i As Integer = 1 To pdfFiles.Count - 1
            stage = "opening " & pdfFiles(i)
            Dim src As Object = CreateObject("AcroExch.PDDoc")
            If Not CBool(src.Open(pdfFiles(i))) Then Throw New Exception("Acrobat could not open " & pdfFiles(i))

            stage = "inserting " & pdfFiles(i)
            Dim lastPage As Integer = CInt(merged.GetNumPages()) - 1
            Dim srcPages As Integer = CInt(src.GetNumPages())
            Dim ok As Boolean = CBool(merged.InsertPages(lastPage, src, 0, srcPages, 0))
            If Not ok Then Throw New Exception("InsertPages returned False")

            src.Close()
        Next

        stage = "saving combined file"
        If Not CBool(merged.Save(1, outPath)) Then Throw New Exception("Save returned False")
        merged.Close()
    Catch ex As Exception
        MessageBox.Show("Merge failed while " & stage & vbCrLf & ex.Message, "PDF export")
        Return
    End Try

    ' 4. Report
    Dim summary As String = "Exported " & pdfFiles.Count & " drawings." & vbCrLf & "Combined file:" & vbCrLf & outPath

    missingAsm.Sort()
    missingParts.Sort()

    Dim report As New System.Text.StringBuilder()
    Dim reportPath As String = ""

    If missingAsm.Count > 0 Then
        report.AppendLine("ASSEMBLIES with no drawing (" & missingAsm.Count & "):")
        For Each nm As String In missingAsm
            report.AppendLine("    " & nm)
        Next
        report.AppendLine()
    End If

    If missingParts.Count > 0 Then
        report.AppendLine("PARTS with no drawing (" & missingParts.Count & "):")
        For Each nm As String In missingParts
            report.AppendLine("    " & nm)
        Next
        report.AppendLine()
    End If

    If skippedFiles.Count > 0 Then
        report.AppendLine("SKIPPED (file locked), not in combined PDF (" & skippedFiles.Count & "):")
        For Each nm As String In skippedFiles
            report.AppendLine("    " & nm)
        Next
    End If

    If report.Length > 0 Then
        reportPath = System.IO.Path.Combine(outFolder, "_Missing drawings.txt")
        System.IO.File.WriteAllText(reportPath, report.ToString())
    End If

    ShowResults(summary, report.ToString().TrimEnd(), reportPath)
End Sub

Function ClearOldFile(filePath As String, allowSkip As Boolean) As Integer
    ' Returns 1 = file is gone (or never existed), 2 = user chose Skip, 0 = user chose Cancel
    While System.IO.File.Exists(filePath)
        Try
            System.IO.File.SetAttributes(filePath, System.IO.FileAttributes.Normal)
            System.IO.File.Delete(filePath)
        Catch ex As Exception
            Dim choice As Integer = AskLocked(filePath, allowSkip)
            If choice = 0 Then Return 0
            If choice = 2 Then Return 2
            ' choice = 1 (Retry): loop round and try again
        End Try
    End While
    Return 1
End Function

Function AskLocked(filePath As String, allowSkip As Boolean) As Integer
    ' Returns 1 = Retry, 2 = Skip, 0 = Cancel
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "PDF export"
    frm.ClientSize = New System.Drawing.Size(460, 150)
    frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False

    Dim lbl As New System.Windows.Forms.Label()
    lbl.AutoSize = False
    lbl.UseMnemonic = False
    lbl.Location = New System.Drawing.Point(12, 12)
    lbl.Size = New System.Drawing.Size(436, 90)
    lbl.Text = "Can't overwrite:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
               "It is probably open in another program. Close it, then click Retry."

    Dim btnRetry As New System.Windows.Forms.Button()
    btnRetry.Text = "Retry"
    btnRetry.Location = New System.Drawing.Point(188, 112)
    btnRetry.Size = New System.Drawing.Size(80, 28)
    btnRetry.DialogResult = System.Windows.Forms.DialogResult.Retry

    Dim btnSkip As New System.Windows.Forms.Button()
    btnSkip.Text = "Skip"
    btnSkip.Location = New System.Drawing.Point(274, 112)
    btnSkip.Size = New System.Drawing.Size(80, 28)
    btnSkip.DialogResult = System.Windows.Forms.DialogResult.Ignore
    btnSkip.Enabled = allowSkip

    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Cancel"
    btnCancel.Location = New System.Drawing.Point(368, 112)
    btnCancel.Size = New System.Drawing.Size(80, 28)
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel

    frm.AcceptButton = btnRetry
    frm.CancelButton = btnCancel
    frm.Controls.Add(lbl)
    frm.Controls.Add(btnRetry)
    frm.Controls.Add(btnSkip)
    frm.Controls.Add(btnCancel)

    Dim result As System.Windows.Forms.DialogResult = frm.ShowDialog()
    frm.Dispose()

    If result = System.Windows.Forms.DialogResult.Retry Then Return 1
    If result = System.Windows.Forms.DialogResult.Ignore Then Return 2
    Return 0
End Function

Sub ShowResults(summary As String, details As String, reportPath As String)
    ' Nothing to report: a plain message box is enough
    If details.Length = 0 Then
        MessageBox.Show(summary, "PDF export")
        Return
    End If

    Dim margin As Integer = 12
    Dim gap As Integer = 8
    Dim btnH As Integer = 28
    Dim clientW As Integer = 520
    Dim innerW As Integer = clientW - 2 * margin
    Dim flags As System.Windows.Forms.TextFormatFlags = _
        System.Windows.Forms.TextFormatFlags.WordBreak Or System.Windows.Forms.TextFormatFlags.NoPrefix

    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "PDF export"
    frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
    frm.MinimizeBox = False

    Dim lbl As New System.Windows.Forms.Label()
    lbl.AutoSize = False
    lbl.UseMnemonic = False
    lbl.Text = summary

    Dim txt As New System.Windows.Forms.TextBox()
    txt.Multiline = True
    txt.ReadOnly = True
    txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    txt.Text = details
    txt.TabStop = False

    ' Measure the text so the layout fits the content
    Dim lblH As Integer = System.Windows.Forms.TextRenderer.MeasureText( _
        summary, lbl.Font, New System.Drawing.Size(innerW, 0), flags).Height + 6

    Dim txtTop As Integer = margin + lblH + gap
    Dim fixedH As Integer = txtTop + margin + btnH + margin   ' everything except the list box

    Dim maxTxtH As Integer = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height - fixedH - 80
    Dim wantTxtH As Integer = System.Windows.Forms.TextRenderer.MeasureText( _
        details, txt.Font, New System.Drawing.Size(innerW - 24, 0), flags).Height + 12
    Dim txtH As Integer = Math.Max(100, Math.Min(wantTxtH, maxTxtH))

    frm.ClientSize = New System.Drawing.Size(clientW, fixedH + txtH)

    lbl.Location = New System.Drawing.Point(margin, margin)
    lbl.Size = New System.Drawing.Size(innerW, lblH)
    txt.Location = New System.Drawing.Point(margin, txtTop)
    txt.Size = New System.Drawing.Size(innerW, txtH)

    Dim btnY As Integer = frm.ClientSize.Height - margin - btnH

    Dim btnOpen As New System.Windows.Forms.Button()
    btnOpen.Text = "Open list in Notepad"
    btnOpen.Location = New System.Drawing.Point(margin, btnY)
    btnOpen.Size = New System.Drawing.Size(140, btnH)
    AddHandler btnOpen.Click, Sub(s, e) System.Diagnostics.Process.Start(reportPath)

    Dim btnOk As New System.Windows.Forms.Button()
    btnOk.Text = "OK"
    btnOk.Location = New System.Drawing.Point(clientW - margin - 80, btnY)
    btnOk.Size = New System.Drawing.Size(80, btnH)
    btnOk.DialogResult = System.Windows.Forms.DialogResult.OK

    ' Anchors: the list box stretches with the window, buttons stick to the bottom
    lbl.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
    txt.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
    btnOpen.Anchor = System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left
    btnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right

    frm.AcceptButton = btnOk
    frm.Controls.Add(lbl)
    frm.Controls.Add(txt)
    frm.Controls.Add(btnOpen)
    frm.Controls.Add(btnOk)

    ' Don't let the window shrink smaller than its starting layout allows
    frm.MinimumSize = New System.Drawing.Size(frm.Width, fixedH + 100 + (frm.Height - frm.ClientSize.Height))

    frm.ShowDialog()
    frm.Dispose()
End Sub
