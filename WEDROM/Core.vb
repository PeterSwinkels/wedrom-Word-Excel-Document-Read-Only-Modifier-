'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   'This structure defines an XML element.
   Private Structure XMLElementStr
      Public Data As String      'Defines the XML element's data.
      Public Offset As Integer   'Defines the XML element's offset. 
   End Structure

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Console.WriteLine(ProgramInformation())
         Console.WriteLine()
         Console.WriteLine(My.Application.Info.Description)

         With New OpenFileDialog With {.CheckFileExists = True, .FileName = Nothing, .Filter = "Microsoft Excel and Word files (*.docx;*.xlsx)|*.docx;*.xlsx", .FilterIndex = 1}
            If Not .ShowDialog() = DialogResult.Cancel Then
               Select Case Path.GetExtension(.FileName).ToLower()
                  Case ".docx"
                     BypassProtection(.FileName, "word\settings.xml", "w:writeProtection")
                  Case ".xlsx"
                     BypassProtection(.FileName, "xl\workbook.xml", "fileSharing")
                  Case Else
                     MessageBox.Show("Unsupported file type!", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
               End Select
            End If
         End With
      Catch ExceptionO As Exception
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Try
   End Sub

   'This procedure bypasses the specified document's protection.
   Private Sub BypassProtection(DocumentPath As String, XMLFile As String, XMLElementName As String)
      Try
         Dim XMLDirectory As String = ExtractXMLFiles(DocumentPath)
         Dim XMLPath As String = Path.Combine(XMLDirectory, XMLFile)
         Dim XMLDocumentContent As String = File.ReadAllText(XMLPath)
         Dim XMLElementText As XMLElementStr = GetXMLElement(XMLElementName, XMLDocumentContent)

         XMLDocumentContent = XMLDocumentContent.Remove(XMLElementText.Offset, XMLElementText.Data.Length)
         File.WriteAllText(XMLPath, XMLDocumentContent)

         ZipFile.CreateFromDirectory(XMLDirectory, DocumentPath)

         OpenDocument(DocumentPath)

         XMLDirectory = ExtractXMLFiles(DocumentPath)

         XMLDocumentContent = File.ReadAllText(XMLPath)
         XMLDocumentContent = XMLDocumentContent.Insert(XMLElementText.Offset, XMLElementText.Data)
         File.WriteAllText(XMLPath, XMLDocumentContent)

         ZipFile.CreateFromDirectory(XMLDirectory, DocumentPath)
         If Directory.Exists(XMLDirectory) Then Directory.Delete(XMLDirectory, recursive:=True)
      Catch ExceptionO As Exception
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Try
   End Sub

   'This procedure extracts the specified Microsoft Office document's XML files to a directory.
   Private Function ExtractXMLFiles(DocumentPath As String) As String
      Dim XMLDirectory As String = Nothing

      Try
         XMLDirectory = $"{DocumentPath}.dir"

         If Directory.Exists(XMLDirectory) Then Directory.Delete(XMLDirectory, recursive:=True)
         ZipFile.ExtractToDirectory(DocumentPath, XMLDirectory)
         File.Delete(DocumentPath)
      Catch ExceptionO As Exception
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Try

      Return XMLDirectory
   End Function

   'This procedure extracts the specified XML element from the specified XML document.
   Private Function GetXMLElement(XMLElementName As String, XMLDocumentContent As String) As XMLElementStr
      Dim XMLElementText As XMLElementStr = Nothing

      Try
         Dim StartPosition As Integer = XMLDocumentContent.IndexOf($"<{XMLElementName}")
         Dim EndPosition As Integer = XMLDocumentContent.IndexOf(">", StartPosition)
         Dim XMLElementData As String = Nothing

         If StartPosition >= 0 AndAlso EndPosition >= 0 Then
            XMLElementData = XMLDocumentContent.Substring(StartPosition, (EndPosition - StartPosition) + 1)
            XMLElementText = New XMLElementStr With {.Data = XMLElementData, .Offset = StartPosition}
         End If
      Catch ExceptionO As Exception
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Try

      Return XMLElementText
   End Function

   'This procedure opens the specified document and waits until it is closed by the user.
   Private Sub OpenDocument(DocumentPath As String)
      Try
         Dim ProcessO As Process = Process.Start(New ProcessStartInfo With {.FileName = DocumentPath, .UseShellExecute = True})

         ProcessO.WaitForExit()
      Catch ExceptionO As Exception
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Try
   End Sub

   'This procedure returns this program's information.
   Private Function ProgramInformation() As String
      Try
         With My.Application.Info
            Return $"{ .AssemblyName} v{ .Version} - by: { .CompanyName}"
         End With
      Catch ExceptionO As Exception
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Try

      Return Nothing
   End Function
End Module
