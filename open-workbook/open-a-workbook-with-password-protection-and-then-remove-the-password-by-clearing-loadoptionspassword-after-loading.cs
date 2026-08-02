// Title: How to Remove Password Protection from an Excel Workbook with Aspose.Cells for .NET
// Description: Shows how to load a password‑protected .xlsx using LoadOptions, clear the workbook password via workbook.Settings.Password = null, and save the file without protection using Aspose.Cells in C#.
// Keywords: Aspose.Cells remove workbook password | C# clear Excel password programmatically | LoadOptions.Password Aspose.Cells | workbook.Settings.Password null | unprotect Excel file .NET | open password protected workbook Aspose | save Excel without password | Excel encryption removal C#
// Common Searches: Aspose.Cells open password protected Excel file | remove password from Excel workbook using C# | clear workbook password after loading with LoadOptions | unprotect Excel file programmatically Aspose.Cells | save unprotected copy of protected workbook .NET
// Developer Intent: Load a protected Excel workbook, strip its password, and write an unprotected version using Aspose.Cells in C#.
// Use Cases: Distribute a confidential spreadsheet as an unprotected copy after internal review. | Batch‑process multiple protected workbooks to remove passwords before data extraction or migration. | Integrate password removal into an automated ETL pipeline that prepares Excel files for downstream systems.
// AI Prompts: Write C# code with Aspose.Cells that opens a password‑protected .xlsx, removes the password, and saves it as a new file. | Explain how to use LoadOptions.Password and workbook.Settings.Password to unprotect an Excel workbook in Aspose.Cells. | Provide error‑handling examples for incorrect passwords when removing protection with Aspose.Cells.

using Aspose.Cells;

// Shows how to load a password‑protected .xlsx using LoadOptions, clear the workbook password via workbook.Settings.Password = null, and save the file without protection using Aspose.Cells in C#.
class RemoveWorkbookPassword
{
    static void Main()
    {
        // Load options with the password of the protected workbook
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "myPassword";

        // Open the password‑protected workbook using the load options
        Workbook workbook = new Workbook("protected.xlsx", loadOptions);

        // Clear the workbook's password to remove protection
        workbook.Settings.Password = null;

        // Save the workbook without password protection
        workbook.Save("unprotected.xlsx");
    }
}
