// Title: Save a refreshed Aspose.Cells pivot table workbook to a specific .xlsx file path in C#
// AI Prompts: Write C# code that builds a workbook, inserts sample data, creates a pivot table, refreshes it, and saves the result to a user‑defined .xlsx file using Aspose.Cells. | Show how to modify the example to save the workbook as a PDF after refreshing the pivot table with Aspose.Cells for .NET. | Demonstrate changing the pivot table source range and directing the saved workbook to a different folder path in C# using Aspose.Cells.
// Common Searches: how to programmatically save an Aspose.Cells workbook after refreshing pivot tables in C# | Aspose.Cells .NET example for creating a pivot table and exporting to a custom Excel file location | C# code to refresh pivot tables before calling Workbook.Save with a specific path | save pivot‑generated workbook to a network share using Aspose.Cells for .NET | Aspose.Cells refresh pivot tables then save workbook as PDF in C#
// Tags: Aspose.Cells workbook.save with custom path | refresh pivot tables Aspose.Cells C# | create pivot table Aspose.Cells worksheet | export workbook to PDF Aspose.Cells .NET | change pivot source range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Illustrates creating a workbook, adding sample data, building and refreshing a pivot table, then saving the workbook to a specified .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Amount";
        cells["A2"].Value = "Food";
        cells["B2"].Value = 1200;
        cells["A3"].Value = "Drink";
        cells["B3"].Value = 800;
        cells["A4"].Value = "Supplies";
        cells["B4"].Value = 450;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: rows and data fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh the pivot table to calculate its data
        sheet.RefreshPivotTables();

        // Specify the output file path
        string outputPath = "ModifiedPivotWorkbook.xlsx";

        // Save the modified workbook (uses Workbook.Save(string) rule)
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook saved to: {outputPath}");
    }
}
