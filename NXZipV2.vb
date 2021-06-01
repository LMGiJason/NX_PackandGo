'This is a modification of the original file from the community.
'Added parameters to allow locating master model drawings in the same folder matching a specific format.
'Added parameters to exclude folder locations from inclusion
'Only minimal testing has been done.

Option Strict Off
Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Environment
Imports System.Collections
Imports NXOpen
Imports NXOpen.UF
Imports System.Text.RegularExpressions

Module NXJournal
    Dim theSession As Session = Session.GetSession()
    Dim theUFSession As UFSession = UFSession.GetUFSession()
    Dim regexMasterModel As Regex
    Dim skipfolders() As String

    Sub Echo(ByVal output As String)
        theSession.ListingWindow.Open()
        theSession.ListingWindow.WriteLine(output)
        theSession.LogFile.WriteLine(output)
    End Sub

    Function GetComponentFullPath(ByVal comp As Assemblies.Component) As String
        Dim partName As String = ""
        Dim refsetName As String = ""
        Dim instanceName As String = ""
        Dim origin(2) As Double
        Dim csysMatrix(8) As Double
        Dim transform(3, 3) As Double
        theUFSession.Assem.AskComponentData(comp.Tag, partName, refsetName,
         instanceName, origin, csysMatrix, transform)

        Return partName
    End Function

    Sub GetComponentsFullPaths(ByVal thisComp As Assemblies.Component, ByVal parts As ArrayList)
        Dim thisPath As String = GetComponentFullPath(thisComp)
        If Not parts.Contains(thisPath) Then parts.Add(thisPath)
        For Each child As Assemblies.Component In thisComp.GetChildren()
            GetComponentsFullPaths(child, parts)
        Next
    End Sub

    Sub Main(ByVal args As String())
        Dim workPart As BasePart = theSession.Parts.Work
        Dim masterModelRegex As String = "" '"[BASENAME]_dwg\d.prt"
        Dim processSessionParts As Boolean = False

        'Do we have a regex for the master drawing?
        If args(0).Contains("[BASENAME]") Then
            masterModelRegex = args(0)
            Echo("Maste Model RegEx=" & masterModelRegex)
        End If
        'Do we have a parameter?
        If args.Length > 1 AndAlso (args(1) = "SESSIONPARTS") Then
            processSessionParts = True
            Echo("Processing all session parts...")
        End If
        'Do we have excludes?
        If args.Length > 2 Then
            skipfolders = args(2).Split("|")
        End If

        If workPart Is Nothing Then
            If args.Length = 0 Then
                Echo("Part file argument expected or work part required")
                Return
            End If

            theSession.Parts.LoadOptions.ComponentsToLoad =
                LoadOptions.LoadComponents.None
            theSession.Parts.LoadOptions.SetInterpartData(True,
                LoadOptions.Parent.All)
            theSession.Parts.LoadOptions.UsePartialLoading = True

            Dim partLoadStatus1 As PartLoadStatus = Nothing
            workPart = theSession.Parts.OpenBaseDisplay(args(0), partLoadStatus1)

            If workPart Is Nothing Then Return
        End If

        Dim paths As ArrayList = New ArrayList
        paths.Add(workPart.FullPath)
        CheckForMasterModel(workPart.FullPath, masterModelRegex, paths)

        Dim root As Assemblies.Component = workPart.ComponentAssembly.RootComponent
        If root IsNot Nothing Then
            GetComponentsFullPaths(root, paths)
        End If

        If processSessionParts Then
            For Each aPart As BasePart In theSession.Parts
                If Not paths.Contains(aPart.FullPath) Then
                    paths.Add(aPart.FullPath)
                    CheckForMasterModel(aPart.FullPath, masterModelRegex, paths)
                End If
            Next
        End If

        Dim tmp_dir As String = "C:\TEMP"
        theUFSession.UF.TranslateVariable("UGII_TMP_DIR", tmp_dir) ' See PR 6271239!

        Dim listFile As String = tmp_dir & "\" & workPart.Leaf & ".txt"
        Echo("List file is: " & listFile)

        Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(listFile, False)
        For Each aPath As String In paths
            If Not ExcludedPath(aPath) Then
                If File.Exists(aPath) Then
                    Echo("Adding: " & aPath)
                    outFile.WriteLine(aPath)
                Else
                    Echo("Not Found: " & aPath)
                End If
            End If
        Next
        outFile.Close()

        Dim outputFolder As String = AskUserOutputPath(tmp_dir)
        Dim archive As String = outputFolder & "\" & workPart.Leaf & ".7z"

        Dim path7z As String = Nothing

        If File.Exists("c:\APPS\7-Zip\7z.exe") Then
            path7z = "c:\APPS\7-Zip\7z"
        ElseIf File.Exists("C:\Program Files (x86)\7-Zip\7z.exe") Then
            path7z = "C:\Program Files (x86)\7-Zip\7z"
        ElseIf File.Exists("C:\Program Files\7-Zip\7z.exe") Then
            path7z = "C:\Program Files\7-Zip\7z"
        Else
            Echo("Cannot find 7-Zip\7z.exe")
            Return
        End If

        Dim command As String = """" & path7z & """ a """ & archive & """ @""" & listFile & """"

        Echo(command)
        Shell(command, , True)

        If File.Exists(archive) Then
            Echo("Created: " & archive)
        Else
            Echo("Failed : " & archive)
        End If

        ' The list file is not deleted in case it needs to be editted and reused
        'File.Delete(listFile)

    End Sub

    Private Function ExcludedPath(filename As String) As Boolean
        If skipfolders Is Nothing Then Return False
        For Each exclude As String In skipfolders
            If filename.StartsWith(exclude, StringComparison.InvariantCultureIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function CheckForMasterModel(ByVal fullpath As String, ByVal match As String, ByRef paths As ArrayList) As Integer
        ' Process the list of files found in the directory. 
        '"[BASENAME]_dwg\d.prt"
        If match.Contains("[BASENAME]") Then
            Dim containingFolder As String = Path.GetDirectoryName(fullpath)
            Dim baseName As String = Path.GetFileNameWithoutExtension(fullpath)
            Dim folderFiles As String() = Directory.GetFiles(containingFolder, "*.prt")
            Dim exactMatch As String = match.Replace("[BASENAME]", baseName)
            For Each fl As String In folderFiles
                If Regex.IsMatch(fl, match, RegexOptions.IgnoreCase) Then
                    If Not paths.Contains(fl) Then
                        paths.Add(fl)
                        Echo("Master!!!: " & fl)
                    End If
                End If
            Next
        End If
    End Function

    Function AskUserOutputPath(strDefaultOutputFolder As String) As String

        Dim strLastPath As String = ""
        Dim strOutputPath As String = ""

        'Key will show up in HKEY_CURRENT_USER\Software\VB and VBA Program Settings
        Try
            'Get the last path used from the registry
            strLastPath = GetSetting("NX journal", "NX Zip", "ExportPath")
            'msgbox("Last Path: " & strLastPath)
        Catch e As ArgumentException
        Catch e As Exception
            MsgBox(e.GetType.ToString)
        Finally
        End Try

        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        ' Then use the following code to create the Dialog window
        ' Change the .SelectedPath property to the default location
        With FolderBrowserDialog1
            ' Desktop is the root folder in the dialog.
            .RootFolder = Environment.SpecialFolder.Desktop
            ' Select the strDefaultOutputFolder directory on entry.
            If Directory.Exists(strLastPath) Then
                .SelectedPath = strLastPath
            Else
                .SelectedPath = strDefaultOutputFolder
            End If
            ' Prompt the user with a custom message.
            .Description = "Select the directory to export .7 zip file"
            If .ShowDialog = DialogResult.OK Then
                ' Display the selected folder if the user clicked on the OK button.
                AskUserOutputPath = .SelectedPath
                ' save the output folder path in the registry for use on next run
                SaveSetting("NX journal", "NX Zip", "ExportPath", .SelectedPath)
            Else
                'user pressed 'cancel', exit the subroutine
                AskUserOutputPath = Nothing
                'exit sub
            End If
        End With

    End Function
End Module

