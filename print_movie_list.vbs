' ----------------------------------------------------------------------
'   Script for printing the contents of a movie directory to a text file.
'   (c) 2019-2020 Markus J. Auvo
'
'   The script prints out the movie contents of a given root directory
'   including its subdirectories into text file. The text file will be
'   placed into the root directory provided by the user, e.g. M:\
'
'   In this case, the movies have been are organized into category folders
'   with names beginning with an underscore _. This script iterates through
'   following video file types: avi, mkv, mp4
'
'	GUIDE:
'
'   The user is prompted for following:
'   - root directory for the movie files
' ----------------------------------------------------------------------

OPTION EXPLICIT

' ----------------------------------------------------------------------
'  Constants and Variables
' ----------------------------------------------------------------------

Const title = "PRINT MOVIE DIRECTORY"

Dim targetTextFile:     targetTextFile = "CINEMA.txt"

Dim strRootDirectory ' Absolute path to folder
Dim intFileCount     ' Counter for movie files

Dim objFSO:         Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objFileToWrite
Dim objROOT
Dim objDirItemList

' ----------------------------------------------------------------------
'  Methods
' ----------------------------------------------------------------------

'
' Displays messages in console window without a line break
'
Sub WriteToConsole(msg)
    WScript.StdOut.Write msg
End Sub

'
' Displays messages in console window without a line break
' and moves the cursor back to the beginning of the line.
'
Sub WriteToConsoleR(msg)
    WScript.StdOut.Write msg & chr(13)
End Sub

'
' Displays messages in console window with a line break
'
Sub WriteToConsoleNL(msg)
    WScript.Echo msg 
End Sub

'
' Displays a message in a dialog box
'
Sub WriteToDialog(msg)
    MsgBox msg, 0, "Message"
End Sub

'
' Writes a line of strings to a file
'
Sub WriteToFile(dataline)
    objFileToWrite.WriteLine(dataline)
End Sub

'
' Gets the root directory at a given path -- if it is found
'
Function getRootDirectory(rootPath)
    If (objFSO.FolderExists(rootPath)) Then
        Set objFileToWrite = objFSO.OpenTextFile(rootPath + targetTextFile, 2, true)
        Set objROOT = objFSO.getFolder(rootPath)
        Set objDirItemList = objROOT.Subfolders
        WriteToConsoleNL "TARGET ROOT DIRECTORY FOUND!"
    Else
        WriteToDialog "TARGET ROOT DIRECTORY NOT FOUND!"
        Wscript.Quit
    End If
End Function

'
' Iterates through folders in a given root directory
'
Function iterateThroughRoot
    Dim objDirItem
    intFileCount = 0

    For Each objDirItem In objDirItemList
        ' If dir name begins with underscore,
        ' it is a movie folder
        If(Left(objDirItem.Name, 1) = "_") Then
            iterateThroughFiles(objDirItem)
            iterateThroughSubfolders(objDirItem)
        End If
    Next
End Function

'
' Iterates through the subfolders in a folder
'
Function iterateThroughSubfolders(parentFolder)
    Dim objDirItem
    Dim subfolder

    For Each objDirItem In parentFolder.Subfolders
        ' There are no more subfolders
        If(objDirItem.Subfolders.count < 1) Then
            iterateThroughFiles(objDirItem)
        ' Find more subfolders
        Else
            iterateThroughSubfolders(objDirItem)
        End If
    Next
End Function

'
' Iterates through files in a subfolder and
' enters them into the text file.
'
Function iterateThroughFiles(subfolder)
    Dim objFile
    Dim strFileExt

    For Each objFile In subfolder.Files
        strFileExt = objFSO.GetExtensionName(objFile.Path)
        ' Count the movie files
        If(strFileExt = "avi" OR strFileExt = "mkv" OR strFileExt = "mp4") Then
            intFileCount = intFileCount + 1
            WriteToFile "(" & intFileCount & ")" & Chr(9) & objFile.Path
            WriteToConsoleR "Items written into file: " & intFileCount
        End If
    Next
End Function

' ----------------------------------------------------------------------
'  Main Processing
' ----------------------------------------------------------------------

'
' User input: Root directory path, by default a drive letter
'
Do While strRootDirectory = Empty
    strRootDirectory = InputBox("Enter the path for movie root", title)
Loop

'
' Do the Magic!!
'
getRootDirectory(strRootDirectory)  ' Gets the root dir
iterateThroughRoot                  ' Iterates through the root dir

'
' Voila!
'
WriteToDialog("Completed!" & Chr(10) & intFileCount & " titles written into file.")
