// Title: Detect Excel format, confirm it is not encrypted, and load the workbook with Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify a spreadsheet's format, checks the IsEncrypted flag, logs the detected LoadFormat, loads the workbook only when it is unencrypted, and prints confirmation messages to the console.
// Keywords: Aspose.Cells FileFormatUtil | detect Excel file format .NET | check workbook encryption C# | load unencrypted workbook Aspose.Cells | Workbook.Settings.IsEncrypted example | C# Excel format detection | Aspose.Cells console logging
// Common Searches: Aspose.Cells detect file format before loading | how to check if Excel file is encrypted using Aspose.Cells | load workbook only when not password protected C# | FileFormatUtil DetectFileFormat example | verify workbook encryption status Aspose.Cells
// Developer Intent: Identify the spreadsheet type, ensure it is not password‑protected, and then open it with Aspose.Cells while providing clear console feedback.
// Use Cases: Validate user‑uploaded Excel files for encryption before processing in a web service. | Log format and encryption details in an automated ETL pipeline that handles multiple spreadsheet sources. | Implement conditional logic based on file type (XLSX, XLS, CSV) only when the file is unencrypted.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an Excel file's format, verify it is not encrypted, and load it with console logging. | Show how to handle an encrypted workbook by prompting for a password and then opening it with Aspose.Cells. | Create a reusable method that returns format, encryption status, and load result for any spreadsheet using Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify a spreadsheet's format, checks the IsEncrypted flag, logs the detected LoadFormat, loads the workbook only when it is unencrypted, and prints confirmation messages to the console.
class Program
{
    static void Main()
    {
        // Path to the workbook file (adjust as needed)
        string filePath = "sample.xlsx";

        // Detect the file format without fully loading the workbook
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Verify that the file is not encrypted
        if (formatInfo.IsEncrypted)
        {
            Console.WriteLine("The workbook is encrypted. Unable to proceed without a password.");
            return;
        }

        // Optionally display the detected load format
        Console.WriteLine($"Detected LoadFormat: {formatInfo.LoadFormat}");

        // Load the unencrypted workbook
        using (Workbook workbook = new Workbook(filePath))
        {
            // Confirm successful loading
            Console.WriteLine("Workbook loaded successfully.");

            // Additional check using workbook settings
            Console.WriteLine($"Workbook Settings IsEncrypted: {workbook.Settings.IsEncrypted}");
        }

        // Final confirmation
        Console.WriteLine("File format verification completed.");
    }
}
