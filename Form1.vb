Public Class frmMain
    Dim Data As XDocument = XDocument.Parse(My.Resources.Resource1.Data)
    Dim UserData As XDocument = XDocument.Parse(My.Resources.Resource1.Data)
    Dim SourcePath As String = System.IO.Directory.GetCurrentDirectory() & "\TagData.xml"
    Dim SaveDirectory As String = System.IO.Directory.GetCurrentDirectory()

    Dim Filename As String = System.IO.Path.GetFileName(SourcePath)
    Dim SavePath As String = System.IO.Path.Combine(SaveDirectory, Filename)
    Dim fade As Image = My.Resources.Resource1.fade
    Dim highlight As Image = My.Resources.Resource1.highlight
    Dim CheckingAll As Boolean = False
    Dim HeroDict As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))
    Dim PortraitDict As Dictionary(Of String, CheckBox) = New Dictionary(Of String, CheckBox)
    Dim SelectColor As Color = Color.Maroon
    Dim DefaultColor As Color = Color.Ivory
    Dim bgColors As Color(,) = New Color(21, 9) {}
    Dim bgDefault As Color(,) = New Color(21, 9) {}

    Function Populate()
        If System.IO.File.Exists(SavePath) Then
            UserData = XDocument.Load(SavePath)
            Return 1
        End If
        Return 0
    End Function

    Function SetupControls()
        For Each button In flpPortraits.Controls.OfType(Of CheckBox)
            AddHandler button.CheckedChanged, AddressOf Portrait_CheckedChanged
        Next
    End Function

    Function MakeDict()
        Dim Heroes As IEnumerable(Of XElement) =
            From el In UserData.<Data>.<Heroes>.<Hero>
            Select el
        For Each Hero In Heroes
            Dim Tags As IEnumerable(Of String) =
                From item In Hero.<Tag>
                Select item.@name
            HeroDict.Add(Hero.@name, Tags.ToList())
        Next
        For Each button In flpPortraits.Controls.OfType(Of CheckBox)
            PortraitDict.Add(button.Tag, button)
        Next
    End Function

    Function RefreshDropDown()
        drpTags.Items.Clear()
        For Each item In chkHero.Items
            drpTags.Items.Add(item)
        Next
    End Function

    Function RefreshChecks()
        For i As Integer = 0 To chkHero.Items.Count() - 1
            chkHero.SetItemChecked(i, False)
        Next
        If HeroDict.ContainsKey(txtHero.Text) Then
            For Each Tag As String In HeroDict(txtHero.Text)
                For i As Integer = 0 To chkHero.Items.Count() - 1
                    If chkHero.Items(i) = Tag Then
                        chkHero.SetItemChecked(i, True)
                    End If
                Next
            Next
            Search()
            PortraitDict(txtHero.Text).Focus()
        End If
    End Function

    Function RefreshAll()
        If Not txtRandom.Text = "" Then
            Dim row = flpPortraits.GetRow(PortraitDict(txtRandom.Text))
            Dim col = flpPortraits.GetColumn(PortraitDict(txtRandom.Text))
            bgColors(col, row) = Color.Transparent
            flpPortraits.Refresh()
        End If
        txtRandom.Clear()
        txtHero.Clear()
        For i As Integer = 0 To chkHero.Items.Count - 1
            chkHero.SetItemChecked(i, False)
        Next
        Search()
    End Function

    Function Search()
        lstSearch.Items.Clear()
        Dim SearchTerms As List(Of String) = New List(Of String)
        For Each item In chkHero.CheckedItems
            SearchTerms.Add(item)
        Next
        For Each Hero As KeyValuePair(Of String, List(Of String)) In HeroDict
            Dim Match As Boolean = True
            For Each item In SearchTerms
                If Not Hero.Value.Contains(item) Then
                    Match = False
                    Exit For
                End If
            Next
            If Match Then
                lstSearch.Items.Add(Hero.Key)
                PortraitDict(Hero.Key).Image = Nothing
            ElseIf chkFade.Checked Then
                PortraitDict(Hero.Key).Image = fade
            End If
        Next
    End Function

    Function RemoveTag()
        For Each item In chkHero.CheckedItems
            Dim ToRemove As IEnumerable(Of XElement) =
                From el In UserData.<Data>.<UserTags>.<Tag>
                Where el.@name = item
                Select el
            For Each Tag As XElement In ToRemove
                Tag.Remove()
            Next
        Next
        For Each item In chkHero.CheckedItems
            Dim ToRemove As IEnumerable(Of XElement) =
                From el In UserData.<Data>.<Heroes>.<Hero>.<Tag>
                Where el.@name = item
                Select el
            For Each Tag As XElement In ToRemove
                Tag.Remove()
            Next
        Next
        chkHero.Items.Clear()
        For Each Tag As XElement In From element In UserData.<Data>.<UserTags>.<Tag>
            chkHero.Items.Add(Tag.@name)
        Next
        Search()
        RefreshDropDown()
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Populate()
        MakeDict()
        SetupControls()
        For Each Tag As XElement In From element In UserData.<Data>.<UserTags>.<Tag>
            chkHero.Items.Add(Tag.@name)
        Next
        Search()
        RefreshDropDown()
    End Sub

    Private Sub Form1_Close(sender As Object, e As EventArgs) Handles MyBase.Closed
        UserData.Save(Filename)
    End Sub

    Private Sub btnSet_Click(sender As Object, e As EventArgs) Handles btnSet.Click
        If Not HeroDict.Keys.Contains(txtHero.Text) Then
            Return
        End If
        Dim NewTags As List(Of String) = New List(Of String)
        Dim Heroes As IEnumerable(Of XElement) =
            From el In UserData.<Data>.<Heroes>.<Hero>
            Where el.@name = txtHero.Text
            Select el
        For Each Hero As XElement In Heroes
            For Each Tag As XElement In Hero.<Hero>
                Tag.Remove()
            Next
            For Each item In chkHero.CheckedItems
                NewTags.Add(item)
                Dim el As XElement = <Tag></Tag>
                el.SetAttributeValue("name", item)
                Hero.Add(el)
            Next
        Next
        HeroDict(txtHero.Text) = NewTags
        Search()
    End Sub

    Private Sub chkHero_SelectedIndexChanged(sender As Object, e As EventArgs) Handles chkHero.SelectedIndexChanged
        chkHero.ClearSelected()
        Search()
    End Sub

    Private Sub txtHero_TextChanged(sender As Object, e As EventArgs) Handles txtHero.TextChanged
        RefreshChecks()
    End Sub

    Private Sub lstSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSearch.SelectedIndexChanged
        txtHero.Text = lstSearch.SelectedItem
    End Sub

    Private Sub btnNewTag_Click(sender As Object, e As EventArgs) Handles btnNewTag.Click
        Dim el As XElement = <Tag></Tag>
        Dim num As Integer = 0
        For Each button In PortraitDict
            If button.Value.Checked Then
                num += 1
            End If
        Next
        Dim NewTag As String = InputBox("Create New Tag for " & num & " selected heroes:", "Create New Tag", "")
        If Not NewTag = "" Then
            el.SetAttributeValue("name", NewTag)
            Dim UserTags As IEnumerable(Of XElement) =
                From Tag In UserData.<Data>.<UserTags>
                Select Tag
            For Each UserTag As XElement In UserTags
                UserTag.Add(el)
                chkHero.Items.Add(el.@name)
            Next
            For Each Hero As XElement In UserData.<Data>.<Heroes>.<Hero>
                If lstSelected.Items.Contains(Hero.@name) Then
                    Hero.Add(el)
                End If
            Next
        End If
        txtHero_TextChanged(txtHero, e)
        RefreshDropDown()
    End Sub

    Private Sub btnRandom_Click(sender As Object, e As EventArgs) Handles btnRandom.Click
        If lstSearch.Items.Count = 0 Then
            Return
        End If
        Dim row = 0
        Dim col = 0
        If Not txtRandom.Text = "" Then
            row = flpPortraits.GetRow(PortraitDict(txtRandom.Text))
            col = flpPortraits.GetColumn(PortraitDict(txtRandom.Text))
            bgColors(col, row) = Color.Transparent
        End If
        txtRandom.Text = lstSearch.Items(New Random().Next(0, lstSearch.Items.Count()))
        row = flpPortraits.GetRow(PortraitDict(txtRandom.Text))
        col = flpPortraits.GetColumn(PortraitDict(txtRandom.Text))
        bgColors(col, row) = Color.Lime
        flpPortraits.Refresh()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        RefreshAll()
    End Sub

    Private Sub btnRemoveTag_Click(sender As Object, e As EventArgs) Handles btnRemoveTag.Click
        Dim answer = MessageBox.Show("Confirm removing all checked tags," & vbCrLf &
         "this cannot be undone.", "Are you sure?", MessageBoxButtons.YesNoCancel)
        If answer = DialogResult.Yes Then
            RemoveTag()
        ElseIf answer = DialogResult.No Then

        End If
    End Sub

    Private Sub btnAll_Click(sender As Object, e As EventArgs) Handles btnAll.Click
        CheckingAll = True
        For Each button In PortraitDict
            button.Value.Checked = True
        Next
        CheckingAll = False
        For Each button In PortraitDict
            lstSelected.Items.Add(button.Key)
        Next
        RefreshAll()
    End Sub

    Private Sub btnNone_Click(sender As Object, e As EventArgs) Handles btnNone.Click
        CheckingAll = True
        For Each button In PortraitDict
            button.Value.Checked = False
        Next
        CheckingAll = False
        lstSelected.Items.Clear()
        RefreshAll()
    End Sub

    Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs)
        lstSelected.Items.Clear()
        If Not CheckingAll Then
            For Each button In PortraitDict
                If button.Value.Checked Then
                    lstSelected.Items.Add(button.Key)
                End If
            Next
        End If
    End Sub

    Private Sub chkFade_CheckedChanged(sender As Object, e As EventArgs) Handles chkFade.CheckedChanged
        If chkFade.Checked Then
            Dim SearchTerms As List(Of String) = New List(Of String)
            For Each item In chkHero.CheckedItems
                SearchTerms.Add(item)
            Next
            For Each Hero As KeyValuePair(Of String, List(Of String)) In HeroDict
                For Each item In SearchTerms
                    If Not Hero.Value.Contains(item) Then
                        PortraitDict(Hero.Key).Image = fade
                        Exit For
                    End If
                Next
            Next
        Else
            For Each Hero As KeyValuePair(Of String, List(Of String)) In HeroDict
                PortraitDict(Hero.Key).Image = Nothing
            Next
        End If
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If Not (drpTags.Text = Nothing) Then
            For Each Hero In lstSelected.Items
                If Not HeroDict(Hero).Contains(drpTags.Text) Then
                    Dim NewTags As List(Of String) = HeroDict(Hero)
                    NewTags.Add(drpTags.Text)
                    HeroDict(Hero) = NewTags
                End If
            Next
            Dim Heroes As IEnumerable(Of XElement) =
                From el In UserData.<Data>.<Heroes>.<Hero>
                Where lstSelected.Items.Contains(el.@name)
                Select el
            For Each Hero As XElement In Heroes
                Dim el As XElement = <Tag></Tag>
                el.SetAttributeValue("name", drpTags.Text)
                Hero.Add(el)
            Next
        End If
        RefreshChecks()
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        If Not (drpTags.Text = Nothing) Then
            For Each Hero In lstSelected.Items
                If HeroDict(Hero).Contains(drpTags.Text) Then
                    HeroDict(Hero).Remove(drpTags.Text)
                End If
            Next
            Dim Tags As IEnumerable(Of XElement) =
                From el In UserData.<Data>.<Heroes>.<Hero>.<Tag>
                Where lstSelected.Items.Contains(el.Parent.@name) And el.@name = drpTags.Text
                Select el
            For Each Tag As XElement In Tags
                Tag.Remove()
            Next
        End If
        RefreshChecks()
    End Sub

    Private Sub Portrait_CheckedChanged(sender As Object, e As EventArgs)
        If Not CheckingAll Then
            If Not My.Computer.Keyboard.CtrlKeyDown Then
                CheckingAll = True
                For Each button As CheckBox In PortraitDict.Values
                    If button IsNot sender Then
                        button.Checked = False
                    End If
                Next
                CheckingAll = False
            End If
            txtHero.Text = sender.Tag
            lstSelected.Items.Clear()
            For Each button In PortraitDict.Values
                If button.Checked Then
                    lstSelected.Items.Add(button.Tag)
                End If
            Next
        End If
    End Sub

    Private Sub Portrait_Click(sender As Object, e As EventArgs)
        If lstSelected.Items.Count = 1 Then
            If lstSelected.Items(0) = sender.Tag Then
                lstSelected.Items.Clear()
                Return
            End If
        End If
        If Not My.Computer.Keyboard.CtrlKeyDown Then
            lstSelected.Items.Clear()
            For Each button As CheckBox In PortraitDict.Values
                button.Checked = False
            Next
        End If
        If lstSelected.Items.Contains(sender.Tag) Then
            lstSelected.Items.Remove(sender.Tag)
        Else
            lstSelected.Items.Add(sender.Tag)
        End If
    End Sub

    Private Sub flpPortraits_CellPaint(sender As Object, e As TableLayoutCellPaintEventArgs) Handles flpPortraits.CellPaint
        Using b As New SolidBrush(bgColors(e.Column, e.Row))
            e.Graphics.FillRectangle(b, e.CellBounds)
        End Using
    End Sub
End Class
