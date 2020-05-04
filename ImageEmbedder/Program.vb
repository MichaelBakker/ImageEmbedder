Imports System.IO
Imports System.Text.RegularExpressions

Module Program

#Region " Objects and variables "


    Private Const EXITME As String = "Druk een toets om af te sluiten."
    Private Const GivePath As String = "Geef het volledige pad op naar de map met afbeeldingen:"
    Private Const PATTERN As String = "bmp|jpg|jpeg|gif|png|svg|webp"
    Private Const PROCESSING As String = "Bezig met verwerken..."
    Private Const WELCOME As String = vbCrLf & "Welkom bij My Kul's Image Embedder!" & vbCrLf & "Geldige afbeelding typen: " & PATTERN

    Private Const fmtBase64 As String = "data:image/{0};base64,{{0}}"
    Private Const fmtCsv As String = "{0}.csv"
    Private Const fmtFileSaved As String = "Excel bestand opgeslagen als '{0}'."
    Private Const fmtFilesFound As String = "{0} bestand(en) gevonden."
    Private Const fmtInvalidFile As String = "Bestand '{0}' overgeslagen: geen geldige afbeelding."
    Private Const fmtNoValidFiles As String = "Geen geldige bestanden gevonden."
    Private Const fmtPattern As String = "$(?<=\.({0}))"
    Private Const fmtRow As String = "{0}{1}{2}"

    Private Const DOT As String = "."
    Private Const BMP As String = "bmp"
    Private Const BMPEXT As String = DOT & BMP
    Private Const JPG As String = "jpg"
    Private Const JPGEXT As String = DOT & JPG
    Private Const JPEG As String = "jpeg"
    Private Const JPEGEXT As String = DOT & JPEG
    Private Const GIF As String = "gif"
    Private Const GIFEXT As String = DOT & GIF
    Private Const PNG As String = "png"
    Private Const PNGEXT As String = DOT & PNG
    Private Const SVG As String = "svg"
    Private Const SVGEXT As String = DOT & SVG
    Private Const WEBP As String = "webp"
    Private Const WEBPEXT As String = DOT & WEBP

    Private ReadOnly fmtBmp As String = String.Format(fmtBase64, BMP)
    Private ReadOnly fmtJpg As String = String.Format(fmtBase64, JPG)
    Private ReadOnly fmtJpeg As String = String.Format(fmtBase64, JPEG)
    Private ReadOnly fmtGif As String = String.Format(fmtBase64, GIF)
    Private ReadOnly fmtPng As String = String.Format(fmtBase64, PNG)
    Private ReadOnly fmtSvg As String = String.Format(fmtBase64, SVG)
    Private ReadOnly fmtWebp As String = String.Format(fmtBase64, WEBP)

    Private ReadOnly searchPattern As New Regex(String.Format(fmtPattern, PATTERN), RegexOptions.IgnoreCase Or RegexOptions.Compiled)

#End Region

    ''' <summary>
    ''' Takes a folder from input and creates a list of embed codes from all the valid images it can find.
    ''' </summary>
    Sub Main()
        Dim strFolder As String = String.Empty
        Dim blFolderInvalid As Boolean = True
        Dim di As DirectoryInfo = Nothing
        Dim files As IEnumerable(Of FileInfo) = Nothing
        Dim data As New List(Of Tuple(Of String, String))
        Dim csvPath As String = String.Empty

        WriteLine(WELCOME)
        WriteLine(vbCrLf)
        Do While String.IsNullOrEmpty(strFolder) AndAlso blFolderInvalid
            WriteLine(GivePath)
            WriteLine(vbCrLf)
            strFolder = ReadLine()
            Try
                di = New DirectoryInfo(strFolder)
                blFolderInvalid = (Not di.Exists)
                csvPath = GetCsvPath(di)
                files = di.EnumerateFiles().Where(Function(f) searchPattern.IsMatch(f.Extension))
            Catch ex As Exception
                WriteLine(ex.Message & vbCrLf & ex.StackTrace, ConsoleColor.Red)
            End Try
        Loop

        If (Not files Is Nothing) AndAlso files.Any Then
            WriteLine(vbCrLf)
            WriteLine(String.Format(fmtFilesFound, files.Count))
            WriteLine(PROCESSING)
            WriteLine(vbCrLf)

            For Each file As FileInfo In files
                Select Case file.Extension
                    Case BMPEXT
                        data.Add(CreateTuple(file, fmtBmp))
                    Case JPGEXT
                        data.Add(CreateTuple(file, fmtJpg))
                    Case JPEGEXT
                        data.Add(CreateTuple(file, fmtJpeg))
                    Case GIFEXT
                        data.Add(CreateTuple(file, fmtGif))
                    Case PNGEXT
                        data.Add(CreateTuple(file, fmtPng))
                    Case SVGEXT
                        data.Add(CreateTuple(file, fmtSvg))
                    Case WEBPEXT
                        data.Add(CreateTuple(file, fmtWebp))
                    Case Else
                        WriteLine(String.Format(fmtInvalidFile, file.Name), ConsoleColor.Red)
                        WriteLine(vbCrLf)
                End Select
            Next

            Try
                CreateCsv(data, csvPath)
                WriteLine(String.Format(fmtFileSaved, csvPath))
            Catch ex As Exception
                WriteLine(ex.Message & vbCrLf & ex.StackTrace, ConsoleColor.Red)
            End Try
        Else
            WriteLine(fmtNoValidFiles, ConsoleColor.Red)
        End If

        WriteLine(vbCrLf)
        WriteLine(EXITME, ConsoleColor.White)
        Console.ReadKey()

    End Sub

    ''' <summary>
    ''' Creates a <see cref="Tuple(Of String, String)"/> from the given image file, containing its name and its base64-encoded image data.
    ''' </summary>
    ''' <param name="file">The image file</param>
    ''' <param name="format">The base64 format</param>
    ''' <returns>A <see cref="Tuple(Of String, String)"/></returns>
    Private Function CreateTuple(file As FileInfo, format As String) As Tuple(Of String, String)

        Return New Tuple(Of String, String)(Path.GetFileNameWithoutExtension(file.FullName), String.Format(format, GetBase64StringForImage(file.FullName)))

    End Function

    ''' <summary>
    ''' Creates a Csv file containing tab separated rows with the image file name and the base64-encoded image data.
    ''' </summary>
    ''' <param name="data">A <see cref="List(Of Tuple(Of String, String))"/></param>
    ''' <param name="csvPath">The path to the csv file to generate</param>
    Private Sub CreateCsv(data As List(Of Tuple(Of String, String)), csvPath As String)

        Using writer As StreamWriter = File.CreateText(csvPath)
            data.ForEach(Sub(t)
                             writer.WriteLine(String.Format(fmtRow, t.Item1, ControlChars.Tab, t.Item2))
                         End Sub)
            writer.Flush()
        End Using

    End Sub

    ''' <summary>
    ''' Reads the file as a byte array and creates a base64 string from that.
    ''' </summary>
    ''' <param name="imgPath">Full path to the file</param>
    Private Function GetBase64StringForImage(imgPath As String) As String

        Return Convert.ToBase64String(File.ReadAllBytes(imgPath))

    End Function

    ''' <summary>
    ''' Returns the path to the created Csv file.
    ''' </summary>
    ''' <param name="folder">The folder from which data was extracted.</param>
    Private Function GetCsvPath(folder As DirectoryInfo) As String

        Return Path.Combine(folder.FullName, String.Format(fmtCsv, folder.Name))

    End Function

    ''' <summary>
    ''' Reads a colored line.
    ''' </summary>
    ''' <param name="color">A <see cref="ConsoleColor"/></param>
    Private Function ReadLine(Optional color As ConsoleColor = ConsoleColor.White) As String

        Console.ForegroundColor = color
        Return Console.ReadLine()

    End Function

    ''' <summary>
    ''' Writes a colored line.
    ''' </summary>
    ''' <param name="text">The text to print out</param>
    ''' <param name="color">A <see cref="ConsoleColor"/></param>
    Private Sub WriteLine(text As String, Optional color As ConsoleColor = ConsoleColor.Green)

        Console.ForegroundColor = color
        Console.WriteLine(text)

    End Sub

End Module