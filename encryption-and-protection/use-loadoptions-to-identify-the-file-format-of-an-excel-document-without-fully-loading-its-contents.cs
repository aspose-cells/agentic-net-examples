// Title: Detect Excel File Format and Build LoadOptions with Aspose.Cells (C#)
// Description: Shows how to use Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, read the file type, encryption and RMS‑protection flags, and then create a LoadOptions instance with the detected LoadFormat—without loading the workbook into memory.
// Keywords: Aspose.Cells DetectFileFormat | FileFormatInfo C# | LoadOptions from file format | Excel encryption detection Aspose | RMS protected Excel detection | Identify Excel load format | Check Excel file type without opening
// Common Searches: Aspose.Cells detect Excel format without loading workbook | How to get LoadFormat of an Excel file in C# | Check if Excel file is encrypted using Aspose.Cells | Create LoadOptions after detecting file type | FileFormatUtil DetectFileFormat example
// Developer Intent: Find the Excel file’s format and security attributes, then instantiate a LoadOptions object with the appropriate LoadFormat, all without opening the workbook.
// Use Cases: Quickly verify the type of large Excel files to avoid unnecessary memory usage. | Determine whether a workbook is password‑protected or RMS‑protected before prompting the user. | Prepare a correctly configured LoadOptions object for subsequent processing with custom load settings.
// AI Prompts: Write C# code that uses Aspose.Cells to detect an Excel file’s format, encryption status, and RMS protection, then returns a LoadOptions object with the detected LoadFormat. | Explain how FileFormatUtil.DetectFileFormat and LoadOptions work together to identify an Excel file type without loading its contents. | Create a reusable method that accepts a file path, returns FileFormatInfo and a pre‑configured LoadOptions, and handles missing file or unsupported format errors.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to use Aspose.Cells.FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object, read the file type, encryption and RMS‑protection flags, and then create a LoadOptions instance with the detected LoadFormat—without loading the workbook into memory.
class DetectExcelFormat
{
    static void Main()
    {
        // Path to the Excel file whose format we want to identify
        string filePath = "sample.xlsx";

        // Detect the file format without loading the entire workbook
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Display detection results
        Console.WriteLine($"File Format Type: {formatInfo.FileFormatType}");
        Console.WriteLine($"Load Format: {formatInfo.LoadFormat}");
        Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
        Console.WriteLine($"Is Protected By RMS: {formatInfo.IsProtectedByRMS}");

        // Create LoadOptions using the detected LoadFormat (no workbook is loaded)
        LoadOptions loadOptions = new LoadOptions(formatInfo.LoadFormat);
        Console.WriteLine($"LoadOptions.LoadFormat: {loadOptions.LoadFormat}");
    }
}
