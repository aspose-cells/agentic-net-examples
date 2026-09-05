// Title: Set the write‑protection author of an Excel workbook to the logged‑in Windows user and save the file using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells, assign Environment.UserName to workbook.Settings.WriteProtection.Author, and save the workbook. | Programmatically change the write‑protection author of a workbook to the current Windows account in C# and overwrite the original file with Aspose.Cells. | Retrieve the current Windows username, set it as the write‑protection author property, and persist the Excel document using Aspose.Cells.
// Common Searches: how to change write protection author to current windows user using Aspose.Cells C# | Aspose.Cells example for setting workbook.Settings.WriteProtection.Author | C# code to assign Environment.UserName as Excel file protection author before saving | update Excel workbook protection metadata with logged in Windows username using Aspose.Cells | overwrite existing .xlsx after modifying write protection author programmatically
// Tags: Excel workbook security settings | dotnet metadata update for Excel | set workbook protection fields programmatically | save Excel workbook with new settings | modify workbook security attributes

using System;
using Aspose.Cells;

// Loads an existing Excel workbook, sets the write‑protection author to the current Windows user via workbook.Settings.WriteProtection.Author, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Change the write‑protection author to the current Windows user name
        workbook.Settings.WriteProtection.Author = Environment.UserName;

        // Re‑save the workbook (overwrite or specify a new file name)
        workbook.Save("output.xlsx");
    }
}
