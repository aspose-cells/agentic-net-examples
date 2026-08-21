// Title: C# – Password‑Protect and Encrypt an Excel Workbook, Save as XLS with Aspose.Cells
// Description: Loads a plain .xlsx file (creates a sample if missing), sets a user‑defined password, applies 128‑bit strong encryption, and saves the workbook in Excel 97‑2003 (XLS) format using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | encrypt Excel workbook | password protect XLS | 128-bit encryption | Workbook.Settings.Password | SetEncryptionOptions | save as XLS | Excel 97-2003 | secure Excel file | convert xlsx to xls
// Common Searches: Aspose.Cells encrypt workbook C# | How to password protect an XLS file using Aspose.Cells | Set 128‑bit encryption for Excel 97‑2003 with Aspose | Convert .xlsx to .xls with password protection .NET | C# code to save encrypted XLS with Aspose.Cells
// Developer Intent: Apply a user‑specified password and strong encryption to an existing workbook and output it as a protected XLS file.
// Use Cases: Securely distribute a financial report by delivering a password‑protected XLS version. | Meet compliance rules that require 128‑bit encryption when converting confidential .xlsx files to the legacy XLS format. | Automate batch processing that reads plain workbooks, encrypts each with a unique password, and stores them as protected XLS files.
// AI Prompts: Show C# code that loads an .xlsx file, sets a user‑provided password, configures 128‑bit strong encryption, and saves the workbook as an XLS file using Aspose.Cells. | Provide an Aspose.Cells example with error handling that encrypts a workbook with a password and saves it in Excel 97‑2003 format.

using System;
using System.IO;
using Aspose.Cells;

// Loads a plain .xlsx file (creates a sample if missing), sets a user‑defined password, applies 128‑bit strong encryption, and saves the workbook in Excel 97‑2003 (XLS) format using Aspose.Cells for .NET.
class EncryptWorkbookToXls
{
    static void Main()
    {
        string inputPath = "plain.xlsx";
        string outputPath = "encrypted.xls";
        string password = "mySecretPassword";

        try
        {
            // Ensure the source workbook exists; create a simple one if it doesn't.
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                Worksheet sheet = tempWb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Set password to protect the workbook.
            workbook.Settings.Password = password;

            // Apply strong encryption (128‑bit key) for XLS format.
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save as Excel 97‑2003 format (XLS).
            workbook.Save(outputPath, SaveFormat.Excel97To2003);

            Console.WriteLine($"Workbook encrypted and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
