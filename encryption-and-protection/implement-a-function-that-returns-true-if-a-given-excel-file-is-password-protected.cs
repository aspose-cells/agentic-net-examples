// Title: Check if an Excel file is password‑protected with Aspose.Cells for .NET (C#)
// Description: A C# helper that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object and returns its IsEncrypted flag, indicating whether the workbook requires a password to open.
// Keywords: Aspose.Cells | C# Excel encryption detection | detect encrypted workbook | FileFormatUtil | FileFormatInfo.IsEncrypted | .NET password protected Excel | Excel file encryption check
// Common Searches: Aspose.Cells check if Excel workbook is encrypted | C# detect password protected Excel file | How to know if an Excel file needs a password using Aspose.Cells | FileFormatUtil DetectFileFormat encrypted workbook
// Developer Intent: Determine whether a specified Excel workbook is encrypted and requires a password before opening.
// Use Cases: Prompt the user for a password only when the file is encrypted, avoiding unnecessary dialogs. | Skip or log encrypted workbooks during bulk import to prevent runtime errors. | Generate an inventory of password‑protected Excel files in a directory for compliance reporting.
// AI Prompts: Write a robust C# method that returns true if an Excel file is password‑protected using Aspose.Cells, including error handling for missing or inaccessible files. | Show how to open an encrypted workbook with Aspose.Cells after confirming protection, prompting the user for the password and handling incorrect entries. | Provide sample code that scans a folder, identifies encrypted Excel files with Aspose.Cells, and writes their full paths to a log file.

using System;
using Aspose.Cells;

// A C# helper that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to obtain a FileFormatInfo object and returns its IsEncrypted flag, indicating whether the workbook requires a password to open.
public class ExcelProtectionChecker
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>True if the file requires a password to open; otherwise, false.</returns>
    public static bool IsFilePasswordProtected(string filePath)
    {
        // Detect the file format and obtain its metadata.
        FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

        // The IsEncrypted property indicates whether the document is encrypted
        // and therefore requires a password to open.
        return fileInfo.IsEncrypted;
    }

    // Example usage
    public static void Main()
    {
        string path = "sample.xlsx";

        bool protectedFlag = IsFilePasswordProtected(path);
        Console.WriteLine($"Is \"{path}\" password protected? {protectedFlag}");
    }
}
