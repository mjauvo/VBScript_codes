' ----------------------------------------------------------------------
'   Script for printing the contents of a movie directory to a text file.
'   (c) 2019-2020 Markus J. Auvo
'
'   VERSION 2
'
'   The script prints out the movie contents of a given root directory,
'   (e.g. M:\) including its subdirectories into an output file (.CSV).
'
'   In this case, the movies have been are organized into category folders
'   with names beginning with an underscore _. This script iterates through
'   following video file types: avi, mkv, mp4
'
'   The output file will be placed into the same directory as this script file.
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
Const outputFile = "CINEMA.csv"

Const MOVIE_FILE = 0
Const MOVIE_SIZE = 1
Const MOVIE_FILE_TYPE = 2
Const MOVIE_YEAR = 15
Const MOVIE_GENRE = 16
Const MOVIE_TAGS = 18
Const MOVIE_TITLE = 21
Const MOVIE_LENGTH = 27
Const MOVIE_FILE_NAME = 165         ' File name including the file extension
Const MOVIE_FOLDER_CURRENT = 190    ' Current folder of the file
Const MOVIE_FOLDER_STRUCTURE =191   ' Folder structure of the file w/o the file
Const MOVIE_FULL_PATH = 194         ' Full path of the file

Dim folderNamespace
Dim strRootDirectory ' Absolute path to folder
Dim intFileCount     ' Counter for movie files

Dim shellAPP:           Set shellAPP = CreateObject("Shell.Application")

Dim objFSO:             Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objOutputFile:      Set objOutputFile = objFSO.CreateTextFile(outputFile, 2, True)

Dim objROOT
Dim objDirItemList

' ----------------------------------------------------------------------
'  Methods
' ----------------------------------------------------------------------

'
' Display messages in console window without a line break
'
Sub WriteToConsole(msg)
    WScript.StdOut.Write msg
End Sub

'
' Display messages without a line break in console window
' and moves the cursor back to the beginning of the line.
'
Sub WriteToConsoleR(msg)
    WScript.StdOut.Write msg & chr(13)
End Sub

'
' Display messages with a line break in console window
'
Sub WriteToConsoleNL(msg)
    WScript.Echo msg 
End Sub

'
' Display a message in a dialog box
'
Sub WriteToDialog(msg)
    MsgBox msg, 0, "Message"
End Sub

'
' Write a line of strings into a file
'
Sub WriteToFile(dataline)
    objOutputFile.WriteLine(dataline)
End Sub

'
' Get the root directory at a given path -- if it is found
'
Function getRootDirectory(rootPath)
    WriteToConsole "Target root directory..."
    If (objFSO.FolderExists(rootPath)) Then
        Set objROOT = objFSO.getFolder(rootPath)
        Set objDirItemList = objROOT.Subfolders
        WriteToConsoleNL "FOUND!"

        Const COL_TITLE = "Title"
        Const COL_YEAR = "Year"
        Const COL_TAGS = "Tags"
        Const COL_LENGTH = "Length"
        Const COL_SIZE = "Size"
        Const COL_FILE_TYPE = "File Type"
        Const COL_FOLDER = "Folder"
        Const COL_PATH = "Path"
        Const COL_FILENAME = "File"

        ' Write column headers into output file
        WriteToFile("sep=;")    ' Make sure that Excel handles the list separator correctly
        WriteToConsole "Column headers for output file..."
        WriteToFile(COL_TITLE & "; " & COL_YEAR & "; " & COL_TAGS & "; " & COL_FOLDER & "; " & COL_LENGTH & "; " & COL_SIZE & "; " & COL_FILE_TYPE & "; " & COL_FILENAME)
        WriteToConsoleNL "OK!"
        WriteToConsoleNL ""

    Else
        WriteToDialog "TARGET ROOT DIRECTORY NOT FOUND!"
        Wscript.Quit
    End If
End Function

'
' Iterate through category folders in a given root directory
'
Function iterateThroughRoot
    Dim objDirItem
    Dim categoryFolderCount: categoryFolderCount = 0
    intFileCount = 0

    For Each objDirItem In objDirItemList
        ' If folder name begins with an underscore,
        ' it is a movie category folder
        If(Left(objDirItem.Name, 1) = "_") Then
            categoryFolderCount = categoryFolderCount + 1
            iterateThroughFiles(objDirItem)
            iterateThroughSubfolders(objDirItem)
        End If
    Next

    ' No category folders were found
    If(categoryFolderCount < 1) Then
        WriteToDialog "NO MOVIE CATEGORIES FOUND!"
    End If
End Function

'
' Iterate through the subfolders in a folder
'
Function iterateThroughSubfolders(parentFolder)
    Dim objDirItem
    Dim subfolder

    For Each objDirItem In parentFolder.Subfolders
        ' There are no more subfolders
        If(objDirItem.Subfolders.count < 1) Then
            iterateThroughFiles(objDirItem)
        ' Go through more subfolders
        Else
            iterateThroughSubfolders(objDirItem)
        End If
    Next
End Function

'
' Iterate through movie files in a subfolder and 
' enter them into the output file.
'
' Movie files of extension AVI, MKV and MP4 are processed.
'
Function iterateThroughFiles(subfolder)
    Dim objFile
    Dim fileName
    Dim strFileExt
    Set folderNamespace = shellAPP.Namespace(subfolder.Path)

    ' Get the movie file properties
    '
    ' This solution was copied and modified from Stackoverflow post reply
    ' Jul 31 '14 at 8:04 by user MC ND.
    ' https://stackoverflow.com/questions/25050807/how-can-i-use-vbscript-to-read-the-attributes-of-an-mp4-file
    ' 
    Dim headers, i, aHeaders(290)
        For i = 0 to 289
            aHeaders(i) = folderNamespace.GetDetailsOf(folderNamespace.Items, i)
        Next

    For Each fileName in folderNamespace.Items
        If (LCase(Right(fileName,4))=".avi" OR LCase(Right(fileName,4))=".mkv" OR LCase(Right(fileName,4))=".mp4") Then 

            Dim movieTitle
            movieTitle = folderNamespace.GetDetailsOf(fileName, MOVIE_TITLE)

                If (len(movieTitle) < 1) Then
                    movieTitle = "#N/A"
                Else
                    movieTitle = chr(34) & movieTitle & chr(34)
                End If

            Dim movieYear
            movieYear = folderNamespace.GetDetailsOf(fileName, MOVIE_YEAR)

                If (len(movieYear) < 1) Then
                    movieYear = "#N/A"
                End If

            Dim movieTags
            movieTags = folderNamespace.GetDetailsOf(fileName, MOVIE_TAGS)

                If (len(movieTags) < 1) Then
                    movieTags = "#N/A"
                End If

            Dim movieLength
            movieLength = folderNamespace.GetDetailsOf(fileName, MOVIE_LENGTH)

                If (len(movieLength) < 1) Then
                    movieLength = "#N/A"
                End If

            Dim movieSize
            movieSize = folderNamespace.GetDetailsOf(fileName, MOVIE_SIZE)

                If (len(movieSize) < 1) Then
                    movieSize = "#N/A"
                End If

            Dim movieFileType
            movieFileType = folderNamespace.GetDetailsOf(fileName, MOVIE_FILE_TYPE)

                If (len(movieFileType) < 1) Then
                    movieFileType = "#N/A"
                Else
                    movieFileType = Left(movieFileType, 3)
                End If

            Dim movieFolderStructure
            movieFolderStructure = folderNamespace.GetDetailsOf(fileName, MOVIE_FOLDER_STRUCTURE)

                If (len(movieFolderStructure) < 1) Then
                    movieFolderStructure = "#N/A"
                Else
                    movieFolderStructure = Right(movieFolderStructure, len(movieFolderStructure)-2)
                    movieFolderStructure = Replace(movieFolderStructure, "\", "  -->  ")
                End If

            Dim movieFile
            movieFile = folderNamespace.GetDetailsOf(fileName, MOVIE_FILE)

                If (len(movieFile) < 1) Then
                    movieFile = "#N/A"
                End If

            ' Write to file
            WriteToFile(movieTitle & "; " movieYear & "; " movieTags & "; " & movieFolderStructure & "; " & movieLength & "; " & movieSize & "; " & movieFileType & "; " & movieFile)

            intFileCount = intFileCount + 1
        End If

        ' Display the operation progress
        WriteToConsoleR "Items written into file: " & intFileCount
    Next
End Function

' ----------------------------------------------------------------------
'  Main Processing
' ----------------------------------------------------------------------

'
' User input: Root directory path, by default a drive root
'
Do While strRootDirectory = Empty
    strRootDirectory = InputBox("Enter the path for movie root", title)
Loop

'
' Do the Magic!!
'
getRootDirectory(strRootDirectory)
iterateThroughRoot

'
' Voila!
'
WriteToDialog("Completed!" & Chr(10) & intFileCount & " titles written into file.")
