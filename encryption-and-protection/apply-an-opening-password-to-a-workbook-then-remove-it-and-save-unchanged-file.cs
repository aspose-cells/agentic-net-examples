// Title: Add an opening password to an Excel workbook, then remove it and save the unchanged file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add opening password to XLSX and later remove it | how to load a password‑protected Excel file and save it without a password using Aspose.Cells | remove workbook password while preserving content with Aspose.Cells in .NET | Aspose.Cells example for clearing opening password from a workbook | save unprotected copy of a protected Excel workbook using C#
// Tags: set workbook opening password Aspose.Cells | clear workbook password Aspose.Cells | load password protected xlsx Aspose.Cells | save workbook without password Aspose.Cells | Aspose.Cells password protection handling C#

using System;
using Aspose.Cells;

// The sample loads a password‑protected XLSX file using LoadOptions with the opening password, clears the workbook's password by setting wb.Settings.Password to null, and saves a new copy without any password, leaving the worksheet data unchanged.
class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";

        // Password used to open the workbook
        string openingPassword = "MySecret";

        // Load the workbook with the opening password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = openingPassword;
        Workbook wb = new Workbook(inputPath, loadOptions);

        // Remove the password protection
        wb.Settings.Password = null; // clearing the password

        // Save the workbook without a password (content unchanged)
        string outputPath = "unprotected.xlsx";
        wb.Save(outputPath);
    }
}
