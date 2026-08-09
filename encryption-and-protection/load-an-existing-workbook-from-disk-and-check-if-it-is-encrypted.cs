// Title: Detect and Load an Encrypted Excel Workbook with Aspose.Cells for .NET
// Description: Shows how to use Aspose.Cells to identify a password‑protected Excel file via FileFormatUtil, load it with LoadOptions when needed, and verify the encryption flag through Workbook.Settings.
// Keywords: Aspose.Cells encryption detection | C# check Excel password protection | FileFormatUtil DetectFileFormat | LoadOptions password protected workbook | Workbook.Settings.IsEncrypted
// Common Searches: how to detect encrypted Excel file using Aspose.Cells | load password protected .xlsx in C# with Aspose | check if workbook is encrypted without opening it | Aspose.Cells get encryption status after loading
// Developer Intent: Identify whether an existing Excel workbook is encrypted and open it with the appropriate password only when required.
// Use Cases: Skip or route protected files in a bulk‑processing pipeline. | Open a password‑protected workbook safely by supplying the password after detection. | Validate that a workbook was opened with correct security settings before performing edits.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect if a .xlsx file is encrypted and open it with a password when necessary. | Explain the role of FileFormatUtil.DetectFileFormat and Workbook.Settings.IsEncrypted in confirming Excel file protection. | Provide error‑handling examples for loading an encrypted workbook with an incorrect password using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to use Aspose.Cells to identify a password‑protected Excel file via FileFormatUtil, load it with LoadOptions when needed, and verify the encryption flag through Workbook.Settings.
class Program
{
    static void Main()
    {
        // Path to the workbook file on disk
        string filePath = "example.xlsx";

        // ------------------------------------------------------------
        // Step 1: Detect encryption without opening the workbook.
        // ------------------------------------------------------------
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        bool isEncrypted = formatInfo.IsEncrypted;
        Console.WriteLine($"Is the workbook encrypted? {isEncrypted}");

        // ------------------------------------------------------------
        // Step 2: Load the workbook.
        // If the file is encrypted, provide the password via LoadOptions.
        // ------------------------------------------------------------
        Workbook workbook;
        if (isEncrypted)
        {
            // Replace "yourPassword" with the actual password for the file.
            LoadOptions loadOptions = new LoadOptions { Password = "yourPassword" };
            workbook = new Workbook(filePath, loadOptions);
        }
        else
        {
            workbook = new Workbook(filePath);
        }

        // ------------------------------------------------------------
        // Step 3: Verify encryption status using the WorkbookSettings property.
        // ------------------------------------------------------------
        Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");
    }
}
