Option Explicit On

Imports System.CodeDom.Compiler
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports Scripting

Public Class Form1
    'Declare public variables
    Public oWord As Microsoft.Office.Interop.Word.Application
    Public srcWord As Microsoft.Office.Interop.Word.Application
    Public srcDoc As Microsoft.Office.Interop.Word.Document
    Public oDoc As Microsoft.Office.Interop.Word.Document
    Public oTable As Microsoft.Office.Interop.Word.Table
    Public oPara As Microsoft.Office.Interop.Word.Paragraph
    Public oPara1 As Microsoft.Office.Interop.Word.Paragraph
    Public oPara2 As Microsoft.Office.Interop.Word.Paragraph
    Public oPara3 As Microsoft.Office.Interop.Word.Paragraph
    Public section As Microsoft.Office.Interop.Word.Section
    Public foot As Microsoft.Office.Interop.Word.Range
    Public lefthead As Microsoft.Office.Interop.Word.Range
    Public centerhead As Microsoft.Office.Interop.Word.Range
    Public righthead As Microsoft.Office.Interop.Word.Range
    Public wordfile As String
    Public pdffile As String
    Public myfilesystemobject As Object
    Public myfiles, myfile, f As Object
    Public pdf_path As String
    Public pdffile_name As String
    Public word_path As String
    Public conv_path As String
    Public convFolder As String
    Public fso As FileSystemObject
    Public row As Integer
    Public wTable As Microsoft.Office.Interop.Word.Table
    Public idx As Integer
    Public idxAlpha As Integer
    Public ProjectDetailsText As String
    Public ClaimantText As String
    Public RespondentText As String

    Sub PDFtoTable()

        'Get the word path
        word_path = TextBox2.Text

        'Create temp folder for converted file
        fso = New FileSystemObject
        convFolder = pdf_path & "\temp"
        conv_path = convFolder

        If fso.FolderExists(convFolder) = False Then
            fso.CreateFolder(convFolder)
        End If

        'PDF conversion to Word
        If ListBox1.Items.Count() > 0 Then
            For Each item In ListBox1.Items

                'Write status to listbox
                StatusBox.AppendText("Converting " & item.ToString() & Environment.NewLine)

                'Get pdf file name
                pdffile_name = Path.GetFileNameWithoutExtension(item)

                'create word document
                srcWord = CreateObject("Word.Application")
                srcWord.Visible = True

                'convert to word
                srcDoc = srcWord.Documents.Open(pdf_path & "\" & pdffile_name & ".pdf")

                'save converted file to temp folder
                srcDoc.SaveAs2(conv_path & "\" & pdffile_name & ".doc")

                srcDoc.Close()

                'Quit converted word document
                srcWord.Quit()
                srcWord = Nothing
                srcDoc = Nothing

                'call sub to create new document
                Call CreateNewDocument()

                'after finish creating doc, save table doc
                oDoc.SaveAs2(word_path & "\" & wordfile & ".doc")

                'Clean up word doc
                oWord.Quit()
                oWord = Nothing
                oDoc = Nothing

                'write progress to listbox
                StatusBox.AppendText("Success!" & Environment.NewLine)


            Next
        End If

        'properly close interop objects
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()

        'kill temp files and folder
        If Directory.Exists(convFolder) Then
            Dim _file As String
            For Each _file In Directory.GetFiles(convFolder)
                Try
                    IO.File.Delete(_file)
                Catch ex As System.IO.IOException
                    StatusBox.AppendText("Unable to delete temp files")
                End Try
            Next
        End If

        'write completed status to listbox after all documents have been converted and applications have been closed properly
        StatusBox.AppendText("Conversion completed!" & Environment.NewLine)

    End Sub

    Sub CreateNewDocument()
        'Create Word and open document template
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add
        Threading.Thread.Sleep(1000)

        'Make this document active
        Dim activeDoc As String
        activeDoc = oWord.ActiveDocument.Name
        oWord.Windows(activeDoc).Activate()

        'Set page orientation to landscape
        oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape

        Call Heading()

    End Sub

    Sub Heading()
        'Insert a paragraph at the beginning of the document
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = tbClaimant.Text & "'s Pleading Review Schedule"
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add
        oPara2.Range.Text = "First Witness Statement of " & tbRespondent.Text
        oPara2.Range.InsertParagraphAfter()

        Call FooterAndHeader()

    End Sub

    Sub FooterAndHeader()
        For Each section In oDoc.Sections
            foot = section.Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
            lefthead = section.Headers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
            centerhead = section.Headers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
            righthead = section.Headers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range

            'Add page number
            With foot
                .Font.Size = 8
                .Font.Bold = True
                .Fields.Add(section.Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range, Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, "Page X of Y", True)
                .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
            End With

            'Add text - Private and Confidential
            With lefthead
                Dim shpCanvas As Microsoft.Office.Interop.Word.Shape
                shpCanvas = section.Headers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 1, 1, 1, 1)
                shpCanvas.TextFrame.TextRange.Text = "Private and Confidential"
                shpCanvas.TextFrame.TextRange.Font.Bold = True
                shpCanvas.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeLeft
                shpCanvas.Top = oWord.InchesToPoints(-0.1)
                shpCanvas.Width = oWord.InchesToPoints(2.5)
                shpCanvas.Height = oWord.InchesToPoints(0.3)
                shpCanvas.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapFront
                shpCanvas.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                shpCanvas.Line.Visible = False

            End With

            'Add text - Company Name
            With centerhead
                Dim shpCanvas As Microsoft.Office.Interop.Word.Shape
                shpCanvas = section.Headers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 1, 1, 1, 1)
                shpCanvas.TextFrame.TextRange.Text = tbProjectDetails.Text
                shpCanvas.TextFrame.TextRange.Font.Bold = True
                shpCanvas.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter
                shpCanvas.Top = oWord.InchesToPoints(-0.1)
                shpCanvas.Width = oWord.InchesToPoints(3.5)
                shpCanvas.Height = oWord.InchesToPoints(0.3)
                shpCanvas.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapFront
                shpCanvas.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                shpCanvas.Line.Visible = False

            End With

            'Add text - date
            With righthead
                Dim shpCanvas As Microsoft.Office.Interop.Word.Shape
                Dim thisDate As String
                thisDate = DateTime.Now.ToString("dd MMMM yyyy")

                shpCanvas = section.Headers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 1, 1, 1, 1)
                shpCanvas.TextFrame.TextRange.Text = thisDate
                shpCanvas.TextFrame.TextRange.Font.Bold = True
                shpCanvas.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeRight
                shpCanvas.Top = oWord.InchesToPoints(-0.1)
                shpCanvas.Width = oWord.InchesToPoints(2.0)
                shpCanvas.Height = oWord.InchesToPoints(0.3)
                shpCanvas.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapFront
                shpCanvas.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                shpCanvas.Line.Visible = False
            End With


        Next

        Call InsertTable()


    End Sub

    Sub InsertTable()
        'Dim variables
        Dim paracount As Integer

        'functions to retrieve data
        paracount = getNumberOfParagraphs()

        'Create table

        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, paracount, 4)
        oTable.Range.ParagraphFormat.SpaceAfter = 0
        oTable.Cell(1, 1).Range.Text = "Paragraph"
        oTable.Cell(1, 2).Range.Text = "Text"
        oTable.Cell(1, 3).Range.Text = "Client's Comment"
        oTable.Cell(1, 4).Range.Text = tbClaimant.Text & "'s Comment"
        oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
        oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle
        oTable.Rows(1).Shading.BackgroundPatternColor = WdColor.wdColorDarkRed
        oTable.Rows(1).Range.Font.ColorIndex = WdColorIndex.wdWhite
        oTable.Columns(1).Width = oWord.InchesToPoints(1.0#)
        oTable.Columns(2).Width = oWord.InchesToPoints(4.3)
        oTable.Columns(3).Width = oWord.InchesToPoints(2.2)
        oTable.Columns(4).Width = oWord.InchesToPoints(2.2)
        For i = 2 To paracount
            oTable.Cell(i, 1).Range.Text = ExtractParagraphs(i)
            oTable.Cell(i, 2).Range.Text = ExtractText(i)
        Next
        oTable.Rows(1).Range.Font.Bold = True

        'apply repeat header rows setting to table
        oWord.ScreenUpdating = False
        oTable.Rows(1).HeadingFormat = False
        oWord.ScreenUpdating = True
        oTable.Rows(1).HeadingFormat = True

        'close src doc
        srcDoc.Close()

        'Quit src doc
        srcWord.Quit()
        srcWord = Nothing
        srcDoc = Nothing

        'CLEAN UP TABLE
        wTable = oDoc.Tables(1)

        'Remove rows if first column is empty and second column contains end of cell markers
        For i = wTable.Rows.Count To 2 Step -1

            'Remove carriage return from cells in table
            Dim rng As Microsoft.Office.Interop.Word.Range
            rng = wTable.Cell(i, 2).Range
            With rng
                .Start = rng.End - 2
            End With
            rng.Text = Replace(rng.Text, vbCr, "")


            Dim strFirstCol As String
            Dim strSecondCol As String

            strFirstCol = wTable.Cell(i, 1).Range.Text
            strSecondCol = wTable.Cell(i, 2).Range.Text

            If Len(wTable.Cell(i, 1).Range.Text) <= 3 And Len(wTable.Cell(i, 2).Range.Text) <= 3 Then
                wTable.Rows(i).Delete()
            End If
        Next i

        'Remove rows if first column is empty and second column contains page number or empty spaces or /
        For i = wTable.Rows.Count To 2 Step -1
            If Len(wTable.Cell(i, 1).Range.Text) <= 3 And Strings.Left(wTable.Cell(i, 2).Range.Text, 1) = Chr(32) Then

                'if cell contain string, then do not delete row
                If CheckForText(wTable.Cell(i, 2).Range.Text) Then
                    wTable.Cell(i, 2).Range.Text = TrimAllWhitespace(wTable.Cell(i, 2).Range.Text)
                    'if cell is empty, then delete row
                ElseIf Not CheckForText(wTable.Cell(i, 2).Range.Text) Then
                    wTable.Rows(i).Delete()

                End If
            End If
        Next i

        'Fix rows with numbered paragraphs in second column if second column has tab
        For i = wTable.Rows.Count To 2 Step -1
            'Trim string
            Dim rng As Microsoft.Office.Interop.Word.Range
            Dim strTrim As String
            Dim textTrim As String

            If Len(wTable.Cell(i, 1).Range.Text) <= 3 And Strings.Left(wTable.Cell(i, 2).Range.Text, 1) = Chr(9) Then

                'Extract numbered list and put in first column
                strTrim = TrimAllWhitespace(wTable.Cell(i, 2).Range.Text)

                'if first char is a number then do this
                If Not CheckForAlphaCharacters(strTrim) Then
                    wTable.Cell(i, 1).Range.Text = Strings.Left(strTrim, idxAlpha)

                    'Extract text from numbered list and put in second column
                    textTrim = Strings.Right(strTrim, Len(strTrim) - idxAlpha)
                    wTable.Cell(i, 2).Range.Text = TrimAllWhitespace(textTrim)

                    'if first char is an alphabet then do this
                ElseIf CheckForAlphaCharacters(strTrim) Then
                    wTable.Cell(i, 1).Range.Text = Strings.Left(strTrim, idxAlpha)

                    'Extract text from numbered list and put in second column
                    textTrim = Strings.Right(strTrim, Len(strTrim) - idxAlpha)
                    wTable.Cell(i, 2).Range.Text = TrimAllWhitespace(textTrim)
                End If
            End If

            'Remove return carriage
            rng = wTable.Cell(i, 2).Range
            With rng
                .Start = rng.End - 2
            End With
            rng.Text = Replace(rng.Text, vbCr, "")
        Next i

        'Fix rows with numbered paragraphs in second column if second column has no tab
        For i = wTable.Rows.Count To 2 Step -1
            'Trim string
            Dim rng As Microsoft.Office.Interop.Word.Range
            Dim textTrim As String
            Dim textBullet As String

            textBullet = Strings.Left(wTable.Cell(i, 2).Range.Text, 4)

            'if range contains bulleted number then do this
            If Len(wTable.Cell(i, 1).Range.Text) <= 3 And IsBulletNumber(textBullet) = True Then

                If textBullet.Contains("Fig.") Then

                Else
                    wTable.Cell(i, 1).Range.Text = Strings.Left(wTable.Cell(i, 2).Range.Text, idx + 1)

                    'Extract text from numbered list and put in second column
                    textTrim = Strings.Right(wTable.Cell(i, 2).Range.Text, Len(wTable.Cell(i, 2).Range.Text) - (idx + 1))
                    wTable.Cell(i, 2).Range.Text = TrimAllWhitespace(textTrim)
                End If
            End If

            'Remove return carriage
            rng = wTable.Cell(i, 2).Range
            With rng
                .Start = rng.End - 2
            End With
            rng.Text = Replace(rng.Text, vbCr, "")
        Next i


        'Fix rows with / due to images or tables converted from pdf
        For i = wTable.Rows.Count To 2 Step -1
            If Len(wTable.Cell(i, 1).Range.Text) <= 3 And Strings.Left(wTable.Cell(i, 2).Range.Text, 1) = Chr(47) Then
                wTable.Rows(i).Delete()
            End If
        Next i



    End Sub


    Function ExtractParagraphs(rowno)
        'Dim variables
        Dim strNumber As String

        wordfile = srcWord.ActiveDocument.Name
        srcWord.Windows(wordfile).Activate()

        strNumber = srcWord.ActiveDocument.Paragraphs(rowno).Range.ListFormat.ListString

        Return strNumber

    End Function

    Function ExtractText(rowno)
        'Dim variables
        Dim strText As String

        wordfile = srcWord.ActiveDocument.Name
        srcWord.Windows(wordfile).Activate()

        strText = srcWord.ActiveDocument.Paragraphs(rowno).Range.Text

        Return strText


    End Function

    Function GetDirPath(ByVal file As String) As String
        Dim f As New FileInfo(file)
        Return f.Directory.ToString
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "Please select a PDF file."
        OpenFileDialog1.Filter = "PDF Files|*.pdf"
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.RestoreDirectory = False

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            ListBox1.Items.Clear()
            OpenFileDialog1.SafeFileNames.Count()

            Dim i As Integer
            For i = 0 To OpenFileDialog1.SafeFileNames.Count() - 1
                ListBox1.Items.Add(OpenFileDialog1.SafeFileNames(i))
            Next

            pdf_path = GetDirPath(OpenFileDialog1.FileName)
            TextBox2.Text = pdf_path

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox2.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If tbProjectDetails.Text = "" And tbClaimant.Text = "" And tbRespondent.Text = "" And ListBox1.Items.Count() <= 0 And TextBox2.Text = "" Then
            MessageBox.Show("You cannot run the tool if there is no input.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf ListBox1.Items.Count() <= 0 And TextBox2.Text = "" Then
            MsgBox("Please select PDF file(s) to run the conversion.")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Please choose a folder to save your file.")
        ElseIf ListBox1.Items.Count() <= 0 Then
            MsgBox("Please select PDF file(s) to run the conversion.")
        ElseIf tbProjectDetails.Text = "" And tbClaimant.Text = "" And tbRespondent.Text = "" Then
            MsgBox("Please provide inputs for the details.")
        ElseIf tbProjectDetails.Text = "" Then
            MsgBox("Please provide project details.")
        ElseIf tbClaimant.Text = "" Then
            MsgBox("Please provide claimant details.")
        ElseIf tbRespondent.Text = "" Then
            MsgBox("Please provide respondent details.")
        Else
            PDFtoTable()

        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ListBox1.Items.Clear()
    End Sub

    Function getNumberOfParagraphs()
        'Dim variables
        Dim oTbl As Microsoft.Office.Interop.Word.Table
        Dim oShp As Microsoft.Office.Interop.Word.Shape
        Dim oPic As Microsoft.Office.Interop.Word.InlineShape
        Dim paracount As Integer

        'Open word doc
        srcWord = CreateObject("word.application")
        srcWord.Visible = False

        'Make opened doc active
        srcDoc = srcWord.Documents.Open(conv_path & "\" & pdffile_name & ".doc")
        wordfile = srcWord.ActiveDocument.Name
        srcWord.Windows(wordfile).Activate()

        'Clean document
        'Delete tables
        For Each oTbl In srcWord.ActiveDocument.Tables
            oTbl.Delete()
        Next oTbl

        'Remove all shapes
        For Each oShp In srcWord.ActiveDocument.Shapes
            oShp.Delete()
        Next oShp

        'Remove pictures
        For Each oPic In srcWord.ActiveDocument.InlineShapes
            oPic.Delete()
        Next oPic

        paracount = srcWord.ActiveDocument.Paragraphs.Count

        Return paracount

    End Function

    Function TrimAllWhitespace(ByVal str As String)

        str = Trim(str)

        Do Until Not Strings.Left(str, 1) = Chr(9)
            str = Trim(Strings.Mid(str, 2, Len(str) - 1))
        Loop

        Do Until Not Strings.Right(str, 1) = Chr(9)
            str = Trim(Strings.Left(str, Len(str) - 1))
        Loop

        Do Until Not Strings.Left(str, 1) = Chr(32)
            str = Trim(Strings.Mid(str, 2, Len(str) - 1))
        Loop

        Do Until Not Strings.Right(str, 1) = Chr(32)
            str = Trim(Strings.Left(str, Len(str) - 1))
        Loop

        TrimAllWhitespace = str

    End Function

    Function CheckForAlphaCharacters(ByVal StringToCheck As String)

        Dim firstChar As String

        firstChar = Strings.Left(StringToCheck, 2)

        For i = 0 To firstChar.Length - 1
            If Char.IsLetter(firstChar.Chars(i)) Then
                idxAlpha = InStr(StringToCheck, ".")
                Return True
            End If
        Next

        Return False

    End Function

    Function CheckForText(ByVal StringToCheck As String)

        For i = 0 To StringToCheck.Length - 1
            If Char.IsLetter(StringToCheck.Chars(i)) Then
                Return True
            End If
        Next

        Return False
    End Function

    Function IsBulletNumber(ByVal NumToCheck As String)
        Dim m As Match

        m = Regex.Match(NumToCheck, "\.")

        If m.Success Then
            idx = InStr(NumToCheck, ".")
            Return True
        End If

        Return False

    End Function





End Class
