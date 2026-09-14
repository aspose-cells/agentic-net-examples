// Title: Decrypt a password‑protected Excel workbook, edit a cell, and re‑encrypt it with a new password using Aspose.Cells for .NET (C#)
// AI Prompts: Show C# code that opens an encrypted .xlsx file with Aspose.Cells, updates cell A1 on the first worksheet, and saves the workbook using a different password. | Explain how to use LoadOptions to load a password‑protected workbook, change a cell value, set Workbook.Settings.Password, and write the file back encrypted with Aspose.Cells.
// Common Searches: C# Aspose.Cells load encrypted xlsx and change cell value | how to change password of an Excel file using Aspose.Cells .NET | update cell in password protected workbook and save with new password Aspose.Cells | Aspose.Cells decrypt workbook, modify data, re‑encrypt in C# example
// Tags: load encrypted xlsx with LoadOptions Aspose.Cells | modify cell value in protected workbook C# | set workbook password using Workbook.Settings Aspose.Cells | re‑encrypt Excel file after editing Aspose.Cells | Aspose.Cells password protection workflow .NET

using System;
using Aspose.Cells;

// Loads a password‑protected XLSX file via LoadOptions, updates cell A1 on the first worksheet, assigns a new password through Workbook.Settings.Password, and saves the workbook encrypted.
class Program
{
    static void Main()
    {
        // Path to the existing password‑protected workbook
        string inputPath = "protected.xlsx";

        // Password used to open the workbook
        string openPassword = "oldPassword";

        // Load the workbook with the password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = openPassword
        };
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Modify a cell value (e.g., Sheet1!A1)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Updated Value");

        // Set a new password for re‑encryption (can be the same as the original)
        string savePassword = "newPassword";
        workbook.Settings.Password = savePassword;

        // Save the workbook with the new encryption
        string outputPath = "protected_modified.xlsx";
        workbook.Save(outputPath);
    }
}
