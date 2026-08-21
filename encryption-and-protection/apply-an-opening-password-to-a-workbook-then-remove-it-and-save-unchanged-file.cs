// Title: Aspose.Cells for .NET – Set and Remove an Opening Password on an Excel Workbook (C#)
// Description: Shows how to assign an opening password to a new workbook via Workbook.Settings.Password, save the protected file, reload it with LoadOptions.Password, clear the password by setting it to null or an empty string, and save the workbook again while keeping all worksheet data unchanged.
// Keywords: Aspose.Cells opening password C# | remove Excel password Aspose | clear workbook password .NET | LoadOptions password protected workbook | Workbook.Settings.Password | Aspose.Cells encryption | programmatically delete Excel password | password‑protected .xlsx handling | Excel file protection C#
// Common Searches: how to add an opening password to an Excel file using Aspose.Cells | remove opening password from a protected workbook Aspose.Cells C# | load password‑protected .xlsx and save without password Aspose | clear workbook password programmatically Aspose.Cells | Aspose.Cells set and clear opening password example
// Developer Intent: Set an opening password on a workbook, then remove it and save the file without altering its content.
// Use Cases: Secure a generated report with a password for external delivery, then strip the password for internal automation. | Protect a template during transport and later open it password‑free for data population. | Batch‑process legacy Excel files that are password‑protected, removing the passwords while preserving original data.
// AI Prompts: Provide C# code that sets an opening password on an Excel workbook with Aspose.Cells, then removes the password and saves the file unchanged. | Show how to load a password‑protected .xlsx using LoadOptions in Aspose.Cells and save it without any opening password. | Explain how to verify that worksheet data remains identical after clearing the opening password in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to assign an opening password to a new workbook via Workbook.Settings.Password, save the protected file, reload it with LoadOptions.Password, clear the password by setting it to null or an empty string, and save the workbook again while keeping all worksheet data unchanged.
class Program
{
    static void Main()
    {
        // ---------- Create a workbook and protect it with an opening password ----------
        Workbook workbook = new Workbook();
        // Add some sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
        // Set the password required to open the file
        workbook.Settings.Password = "open123";
        // Save the password‑protected workbook
        workbook.Save("protected.xlsx");

        // ---------- Load the protected workbook using the password ----------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "open123";
        Workbook loadedWorkbook = new Workbook("protected.xlsx", loadOptions);

        // ---------- Remove the opening password ----------
        // Setting the password to null (or empty string) clears the protection
        loadedWorkbook.Settings.Password = null;

        // Save the workbook back, overwriting the original file (content unchanged)
        loadedWorkbook.Save("protected.xlsx");
    }
}
