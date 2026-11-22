// Title: Set the Title built‑in property of an Excel workbook with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells set workbook built‑in Title property | How to assign a project name to Excel file Title using Aspose.Cells | Programmatically update Excel built‑in document properties with Aspose.Cells .NET | Saving Excel workbook with custom Title metadata in C# Aspose.Cells | Aspose.Cells example for setting built‑in document property Title
// Tags: Aspose.Cells built‑in document property assignment | C# set Excel workbook Title metadata | Aspose.Cells workbook built‑in Title example | Excel file custom Title property using Aspose.Cells | programmatic Excel metadata update Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook with Aspose.Cells, sets the built‑in Title property to "Descriptive Project Name" via the BuiltInDocumentProperties indexer, saves the file as ProjectWorkbook.xlsx, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the Title built‑in property to a descriptive project name
            // The indexer returns a BuiltInDocumentProperty whose Value can be set
            workbook.BuiltInDocumentProperties["Title"].Value = "Descriptive Project Name";

            // Save the workbook to a file
            workbook.Save("ProjectWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
