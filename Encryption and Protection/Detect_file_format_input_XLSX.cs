using System;
using Aspose.Cells;

class DetectFileFormatDemo
{
    static void Main()
    {
        // Path to the input XLSX file
        string filePath = "input.xlsx";

        // Detect the file format using Aspose.Cells utility
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Display detection results
        Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
        Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
    }
}