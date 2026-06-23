using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (empty)
        Workbook originalWorkbook = new Workbook();

        // Save the workbook to a memory stream in XLSX format
        using (MemoryStream memoryStream = new MemoryStream())
        {
            originalWorkbook.Save(memoryStream, SaveFormat.Xlsx);
            // Reset stream position for reading
            memoryStream.Position = 0;

            // Detect file format and encryption status from the stream
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(memoryStream);
            Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted (FileFormatInfo): {formatInfo.IsEncrypted}");

            // Reset stream position again before loading the workbook
            memoryStream.Position = 0;

            // Load the workbook from the memory stream
            Workbook loadedWorkbook = new Workbook(memoryStream);

            // Output encryption status using the loaded workbook's settings
            Console.WriteLine($"Is Encrypted (Workbook.Settings): {loadedWorkbook.Settings.IsEncrypted}");
        }
    }
}