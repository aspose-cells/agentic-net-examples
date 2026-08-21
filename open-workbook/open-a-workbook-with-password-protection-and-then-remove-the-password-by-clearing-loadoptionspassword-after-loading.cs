// Title: C# – Remove password protection from an Excel workbook using Aspose.Cells for .NET
// Description: Load a password‑protected Excel file with Aspose.Cells LoadOptions.Password, clear the workbook password via workbook.Settings.Password, and save the file as an unprotected workbook.
// Keywords: Aspose.Cells C# remove password | LoadOptions.Password Aspose.Cells | workbook.Settings.Password clear | open password protected Excel .NET | save unprotected workbook Aspose | remove Excel workbook password C#
// Common Searches: How to open a password‑protected Excel file with Aspose.Cells and save it without a password | Aspose.Cells .NET clear workbook password after loading | Remove password from .xlsx using LoadOptions.Password in C# | Programmatically delete workbook password with Aspose.Cells
// Developer Intent: Open a password‑protected Excel workbook and save a copy that has no password.
// Use Cases: Strip passwords from confidential spreadsheets before sharing them with a broader audience. | Batch‑process multiple protected workbooks to remove passwords prior to data extraction or analysis. | Integrate into migration pipelines that convert secured Excel files to unprotected versions for downstream systems.
// AI Prompts: Generate C# code using Aspose.Cells to open a password‑protected workbook, remove its password, and save it as a new file. | Explain how to clear workbook.Settings.Password after loading a workbook with LoadOptions.Password in Aspose.Cells. | Show how to handle incorrect password exceptions while removing workbook protection in C#.

using System;
using Aspose.Cells;

// Load a password‑protected Excel file with Aspose.Cells LoadOptions.Password, clear the workbook password via workbook.Settings.Password, and save the file as an unprotected workbook.
class RemoveWorkbookPassword
{
    static void Main()
    {
        // Specify the password used to protect the workbook
        string workbookPassword = "test";

        // Create LoadOptions and set the password to open the protected file
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = workbookPassword;

        // Load the password‑protected workbook using the load options
        Workbook workbook = new Workbook("protected.xlsx", loadOptions);

        // Clear the workbook password to remove protection
        workbook.Settings.Password = null;

        // Save the workbook without password protection
        workbook.Save("unprotected.xlsx");
    }
}
