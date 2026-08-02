// Title: C# Aspose.Cells – Save Workbook with SUM Formula to XLSX (Watch Window Not Available)
// Description: Creates a new workbook, writes values to A1 and B2, adds a SUM formula in C1, notes the lack of a WatchWindow API, and saves the file as WatchWindowDemo.xlsx in XLSX format with error handling.
// Keywords: Aspose.Cells C# save workbook | Aspose.Cells export to XLSX | Aspose.Cells add formula | Aspose.Cells watch window limitation | .NET spreadsheet automation | create Excel file with formula C#
// Common Searches: how to save an Aspose.Cells workbook in C# | Aspose.Cells add SUM formula and export | watch window support in Aspose.Cells .NET | save workbook after configuring watch window Aspose | Aspose.Cells C# example save to xlsx
// Developer Intent: Generate an Excel file with a formula using Aspose.Cells in C# and persist it, while acknowledging that a Watch Window feature is not provided.
// Use Cases: Automate creation of a calculation sheet with sample data and a SUM formula for downstream processing. | Demonstrate workbook saving workflow when the Watch Window API is unavailable. | Provide a minimal reproducible example for debugging formulas in Excel after export.
// AI Prompts: Write C# code with Aspose.Cells to add values, set a SUM formula, and save as XLSX. | Explain alternatives to the missing WatchWindow API for inspecting cell values during development. | Give a step‑by‑step guide to create a workbook, apply a formula, handle exceptions, and save the file.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, writes values to A1 and B2, adds a SUM formula in C1, notes the lack of a WatchWindow API, and saves the file as WatchWindowDemo.xlsx in XLSX format with error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].Formula = "=SUM(A1:B1)";

            // Aspose.Cells does not expose a WatchWindow API; related code is omitted.

            // Define output file path
            string outputPath = "WatchWindowDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
