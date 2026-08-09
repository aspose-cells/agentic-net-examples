// Title: Set Write‑Protection Author to Current Windows User and Overwrite Workbook with Aspose.Cells for .NET
// Description: Loads an existing Excel file, assigns the current Windows user (Environment.UserName) as the write‑protection author, and saves the workbook back to the original location, overwriting the file.
// Keywords: Aspose.Cells | .NET | C# | write protection author | Environment.UserName | overwrite Excel workbook | update workbook protection
// Common Searches: Aspose.Cells change write protection author C# | set Excel write protection author to Windows user | overwrite Excel file after updating protection settings | update workbook protection metadata programmatically
// Developer Intent: Update the workbook's write‑protection author to the logged‑in Windows user and save the changes in place.
// Use Cases: Ensure the author of write‑protected Excel files reflects the person who last processed the document. | Create an audit trail by embedding the current Windows username into the protection metadata of each workbook. | Run a batch operation that updates and overwrites multiple workbooks without generating duplicate files.
// AI Prompts: Provide C# code that sets workbook.Settings.WriteProtection.Author to Environment.UserName using Aspose.Cells and saves the file in place. | Show an example of updating the write‑protection author of an existing Excel workbook and overwriting the original file with Aspose.Cells for .NET. | Explain how to programmatically change the write‑protection author of a workbook, preserve other settings, and save without creating a new copy.

using System;
using Aspose.Cells;

// Loads an existing Excel file, assigns the current Windows user (Environment.UserName) as the write‑protection author, and saves the workbook back to the original location, overwriting the file.
class Program
{
    static void Main()
    {
        // Path to the workbook that needs its write‑protection author updated
        string workbookPath = "input.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(workbookPath);

        // Update the write‑protection author to the current Windows user name
        workbook.Settings.WriteProtection.Author = Environment.UserName;

        // Re‑save the workbook (overwrites the original file)
        workbook.Save(workbookPath);
    }
}
