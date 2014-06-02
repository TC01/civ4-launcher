'Civilization IV Mod Launcher by TC01
'Launches mods for Civ 4, its expansions, and Colonization
'Tries to get the install dir from the registry, if this fails, prompts the user
'Writes the install dir to My.Settings, along with whether or not all versions of Civ are installed
'Does NOT support My Documents\My Games mods.

'As of Version 2, supports Gold and Complete installation paths

Public Class Form1

    Private CivilizationPath, ColonizationPath, WarlordsPath, BTSPath As String
    Private RegRoot As String = "Firaxis Games\"
    Private LocalMachine As String = "HKEY_LOCAL_MACHINE"
    Private CurrentUser As String = "HKEY_CURRENT_USER"
    Private GoldKey As String = "Sid Meier's Civilization 4 Gold"
    Private CompleteKey As String = "Sid Meier's Civilization 4 Complete"

    'Form load function. Used to get registry paths, if registry paths fail, provide prompt for manual install. Also exclude Civ versions not present on local machine.
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim NoCiv As Boolean = My.Settings.NoCiv
        Dim NoWarlords As Boolean = My.Settings.NoWarlords
        Dim NoBTS As Boolean = My.Settings.NoBTS
        Dim NoCol As Boolean = My.Settings.NoCol
        Dim CivInputPath As String = My.Settings.OtherCivPath
        Dim WarlordsInputPath As String = My.Settings.OtherWarlordsPath
        Dim BTSInputPath As String = My.Settings.OtherBTSPath
        Dim ColInputPath As String = My.Settings.OtherColPath

        'Check for Civ 4 first
        If NoCiv = False Then
            If CivInputPath = "" Then
                CivilizationPath = FindInstallPath("", "Sid Meier's Civilization 4", "Mod Launcher has failed to detect your Civ install path. If Civ 4 is installed, find your install path manually and place it in the prompt below. If Civ 4 is not installed, press OK and continue.", "Civ 4 is not installed.")
                If CivilizationPath = "" Then
                    My.Settings.NoCiv = True
                    My.Settings.NoWarlords = True
                    My.Settings.NoBTS = True
                    NoCiv = True
                    NoWarlords = True
                    NoBTS = True
                Else
                    My.Settings.OtherCivPath = CivilizationPath
                End If
            Else
                CivilizationPath = CivInputPath & "\"
            End If

            'If we found Civ 4, keep checking for the expansions (warlords and BTS)
            If NoWarlords = False Then
                If WarlordsInputPath = "" Then
                    WarlordsPath = FindInstallPath("\Warlords", "Sid Meier's Civilization 4 - Warlords", "Mod Launcher has failed to detect your Warlords install path. If Warlords is installed, find your install path manually and place it in the prompt below. If Warlords is not installed, press OK and continue.", "Warlords is not installed.")
                    If WarlordsPath = "" Then
                        My.Settings.NoWarlords = True
                        NoWarlords = True
                    Else
                        My.Settings.OtherWarlordsPath = WarlordsPath
                    End If
                Else
                    WarlordsPath = WarlordsInputPath & "\"
                End If
            End If
            If NoBTS = False Then
                If BTSInputPath = "" Then
                    BTSPath = FindInstallPath("\Beyond the Sword", "Sid Meier's Civilization 4 - Beyond the Sword", "Mod Launcher has failed to detect your Beyond the Sword install path. If Beyond the Sword is installed, find your install path manually and place it in the prompt below. If Beyond the Sword is not installed, press OK and continue.", "Beyond the Sword is not installed.")
                    If BTSPath = "" Then
                        My.Settings.NoBTS = True
                        NoBTS = True
                    Else
                        My.Settings.OtherBTSPath = BTSPath
                    End If
                Else
                    BTSPath = BTSInputPath & "\"
                End If
            End If
        End If

        'Now, check for Colonization (no dependencies)
        If NoCol = False Then
            If ColInputPath = "" Then
                ColonizationPath = FindInstallPath("\Colonization", "Sid Meier's Civilization IV Colonization", "Mod Launcher has failed to detect your Colonization install path. If Colonization is installed, find your install path manually and place it in the prompt below. If Colonization is not installed, press OK and continue.", "Colonization is not installed.")
                If ColonizationPath = "" Then
                    My.Settings.NoCol = True
                    NoCol = True
                Else
                    My.Settings.OtherColPath = ColonizationPath
                End If
            Else
                ColonizationPath = ColInputPath & "\"
            End If
        End If

        'Remove the absent versions from the Combo Box
        If NoCiv = True Then
            ComboBox1.Items.Remove("Civilization 4 Vanilla")
        End If
        If NoWarlords = True Then
            ComboBox1.Items.Remove("Warlords")
        End If
        If NoBTS = True Then
            ComboBox1.Items.Remove("Beyond the Sword")
        End If
        If NoCol = True Then
            ComboBox1.Items.Remove("Civilization IV Colonization")
        End If

        'Set checkbox settings
        CheckBox1.Checked = My.Settings.CloseOnLaunch
        CheckBox2.Checked = My.Settings.Remember
        If My.Settings.Remember = True Then
            ComboBox1.Text = My.Settings.LastGameLaunched
            ComboBox2.Text = My.Settings.LastModLaunched
        End If

        Randomize()

    End Sub

    'Called on combo box selection change. Updates list of available mods per different Civ version.
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim CivVersion As String = ComboBox1.GetItemText(ComboBox1.SelectedItem())
        Dim ModsList As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        Dim CivModString As String
        Dim CivModStart, OldModsList, Count As Integer
        OldModsList = ComboBox2.Items.Count()
        ComboBox2.Items.Clear()

        ' Get the list of mods (default vanilla)
        If Not My.Settings.NoCiv Then
            ModsList = My.Computer.FileSystem.GetDirectories(CivilizationPath & "Mods\")
            If CivVersion = "Warlords" Then
                ModsList = My.Computer.FileSystem.GetDirectories(WarlordsPath & "Mods\")
            ElseIf CivVersion = "Beyond the Sword" Then
                ModsList = My.Computer.FileSystem.GetDirectories(BTSPath & "Mods\")
            End If
        End If

        ' Then do Colonization as a special case.
        If CivVersion = "Civilization IV Colonization" Then
            ' The Colonization mod list may not exist...
            If My.Computer.FileSystem.DirectoryExists(ColonizationPath & "Mods\") Then
                ModsList = My.Computer.FileSystem.GetDirectories(ColonizationPath & "Mods\")
                Count = ModsList.Count()
            Else
                Count = 0
            End If
        End If

        ComboBox2.Items.Add("Regular " & CivVersion)        'So we can start up unmodded Civ versions too!
        For CivMod As Integer = 0 To Count - 1 Step 1
            CivModString = ModsList.Item(CivMod)
            CivModStart = CivModString.LastIndexOf("\")
            CivModString = CivModString.Substring(CivModStart + 1)
            ComboBox2.Items.Add(CivModString)
        Next

    End Sub

    'Called when the mod launch button is clicked. Uses Shell() to launch the given Civ version and mod version.
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim CivVersion As String = ComboBox1.GetItemText(ComboBox1.SelectedItem())
        Dim ModName As String = ComboBox2.GetItemText(ComboBox2.SelectedItem())
        'Do the actual running of the program: through Shell()
        If CivVersion = "Civilization 4 Vanilla" Then
            If ModName = "Regular Civilization 4 Vanilla" Then
                LaunchMod("""" & CivilizationPath & "Civilization4.exe""", ModName)
            Else
                LaunchMod("""" & CivilizationPath & "Civilization4.exe"" mod=\" & ModName, ModName)
            End If
        ElseIf CivVersion = "Warlords" Then
            If ModName = "Regular Warlords" Then
                LaunchMod("""" & WarlordsPath & "Civ4Warlords.exe""", ModName)
            Else
                LaunchMod("""" & WarlordsPath & "Civ4Warlords.exe"" mod=\" & ModName, ModName)
            End If
        ElseIf CivVersion = "Beyond the Sword" Then
            If ModName = "Regular Beyond the Sword" Then
                LaunchMod("""" & BTSPath & "\Civ4BeyondSword.exe""", ModName)
            Else
                LaunchMod("""" & BTSPath & "\Civ4BeyondSword.exe"" mod=\" & ModName, ModName)
            End If
        ElseIf CivVersion = "Civilization IV Colonization" Then
            If ModName = "Regular Civilization IV Colonization" Then
                LaunchMod("""" & ColonizationPath & "\Colonization.exe""", ModName)
            Else
                LaunchMod("""" & ColonizationPath & "\Colonization.exe"" mod=\" & ModName, ModName)
            End If
        End If
    End Sub

    'Called when the reset all button is clicked, resets all settings variables to False (for bools) and "" (for strings)
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        My.Settings.NoCiv = False
        My.Settings.NoWarlords = False
        My.Settings.NoBTS = False
        My.Settings.NoCol = False
        My.Settings.OtherCivPath = ""
        My.Settings.OtherWarlordsPath = ""
        My.Settings.OtherBTSPath = ""
        My.Settings.OtherColPath = ""
        My.Settings.CloseOnLaunch = False
        My.Settings.Remember = False
        My.Settings.LastModLaunched = ""
        My.Settings.LastGameLaunched = ""
        MsgBox("All settings have been cleared!", MsgBoxStyle.OkOnly, "Mod Launcher")
    End Sub

    'Launch when the random mod button is clicked
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim CivVersion As String = ComboBox1.GetItemText(ComboBox1.SelectedItem())
        Dim ModName As String
        Dim ModNum As Integer
        Dim Result As MsgBoxResult

        'Fix a possible unhandled exception
        If CivVersion = Nothing Then
            Return
        End If

        'Get a random modm set a combo box, prompt the user
        ModNum = CInt(Int(((ComboBox2.Items.Count() - 1) - 0 + 1) * Rnd() + 0))
        ModName = ComboBox2.Items.Item(ModNum)
        ComboBox2.Text = ModName
        Result = MsgBox("Launching " & ModName & " as random mod.", vbOKCancel, "Mod Launcher")
        If Result = vbCancel Then
            Return
        End If

        'Do the actual running of the program: through Shell()
        If CivVersion = "Civilization 4 Vanilla" Then
            If ModName = "Regular Civilization 4 Vanilla" Then
                LaunchMod("""" & CivilizationPath & "Civilization4.exe""", ModName)
            Else
                LaunchMod("""" & CivilizationPath & "Civilization4.exe"" mod=\" & ModName, ModName)
            End If
        ElseIf CivVersion = "Warlords" Then
            If ModName = "Regular Warlords" Then
                LaunchMod("""" & WarlordsPath & "Civ4Warlords.exe""", ModName)
            Else
                LaunchMod("""" & WarlordsPath & "Civ4Warlords.exe"" mod=\" & ModName, ModName)
            End If
        ElseIf CivVersion = "Beyond the Sword" Then
            If ModName = "Regular Beyond the Sword" Then
                LaunchMod("""" & BTSPath & "\Civ4BeyondSword.exe""", ModName)
            Else
                LaunchMod("""" & BTSPath & "\Civ4BeyondSword.exe"" mod=\" & ModName, ModName)
            End If
        ElseIf CivVersion = "Civilization IV Colonization" Then
            If ModName = "Regular Civilization IV Colonization" Then
                LaunchMod("""" & ColonizationPath & "\Colonization.exe""", ModName)
            Else
                LaunchMod("""" & ColonizationPath & "\Colonization.exe"" mod=\" & ModName, ModName)
            End If
        End If
    End Sub

    'Change settings for last mod launched
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        My.Settings.CloseOnLaunch = CheckBox1.Checked
    End Sub

    'Change settings for whether or not to save launch data
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        My.Settings.Remember = CheckBox2.Checked
    End Sub

    'Might be necessary, might not be... but just to be sure, saving settings
    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        My.Settings.Save()
    End Sub

    'End of Form1 events, beginning of internal functions and subroutines

    'Finds the path of the given civ version (reg key) passed to it
    Private Function FindInstallPath(ByVal CivType As String, ByVal RegPath As String, ByVal ErrorMsg As String, ByVal DefaultResponse As String)
        Dim Valid As Boolean = True
        Dim Path As String = ""
        Dim DialogPath As String = ""
        Dim RegKey As Object
        Dim RCurrentUser, RLocalMachine As String
        Dim OSType As String

        'Bugfix to check for x64 systems
        If IntPtr.Size = 8 Then
            OSType = "\SOFTWARE\Wow6432Node\"
        Else
            OSType = "\Software\"
        End If

        'Construct the root of both registry strings
        RLocalMachine = LocalMachine & OSType & RegRoot
        RCurrentUser = CurrentUser & OSType & RegRoot

        'First standard, gold, and complete registries on local machine
        RegKey = My.Computer.Registry.GetValue(RLocalMachine & RegPath, "INSTALLDIR", Nothing)
        If CStr(RegKey) = "" Then
            Valid = False
            RegKey = My.Computer.Registry.GetValue(RLocalMachine & GoldKey, "INSTALLDIR" & CivType, Nothing)
            If CStr(RegKey) = "" Then
                RegKey = My.Computer.Registry.GetValue(RLocalMachine & CompleteKey, "INSTALLDIR" & CivType, Nothing)
                If CStr(RegKey) <> "" Then
                    Valid = True
                End If
            Else
                Valid = True
            End If
        End If

        'Then try registries on current user if the others failed
        If Valid = False Then
            RegKey = My.Computer.Registry.GetValue(RCurrentUser & RegPath, "INSTALLDIR", Nothing)
            If CStr(RegKey) = "" Then
                RegKey = My.Computer.Registry.GetValue(RCurrentUser & GoldKey, "INSTALLDIR" & CivType, Nothing)
                If CStr(RegKey) = "" Then
                    RegKey = My.Computer.Registry.GetValue(RCurrentUser & CompleteKey, "INSTALLDIR" & CivType, Nothing)
                    If CStr(RegKey) <> "" Then
                        Valid = True
                    End If
                Else
                    Valid = True
                End If
            Else
                Valid = True
            End If
        End If

        'If nothing worked, make an input box, and save either the registry result or the input result to path
        If Valid = False Then
            DialogPath = InputBox(ErrorMsg, "Mod Launcher", DefaultResponse)
            If DialogPath <> DefaultResponse Then
                Path = DialogPath & "\"
            End If
        Else
            Path = CStr(RegKey) & "\"
        End If

        Return Path
    End Function

    'Launches a mod and perform various special tasks
    Private Sub LaunchMod(ByVal Command As String, ByVal ModName As String)
        Shell(Command)
        If My.Settings.Remember = True Then
            My.Settings.LastGameLaunched = ComboBox1.GetItemText(ComboBox1.SelectedItem)
            My.Settings.LastModLaunched = ModName
        End If
        If My.Settings.CloseOnLaunch = True Then
            Me.Close()
        End If
    End Sub

End Class
