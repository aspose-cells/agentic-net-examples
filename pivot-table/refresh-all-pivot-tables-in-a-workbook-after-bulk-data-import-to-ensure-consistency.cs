// Title: Refresh every PivotTable in an Excel workbook after bulk data import using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing Excel file or create a new workbook, populate a large range of cells, call workbook.Worksheets.RefreshPivotTables(), and save the updated file with Aspose.Cells. | Insert thousands of rows into a worksheet, invoke the RefreshPivotTables method on all worksheets to synchronize pivot tables, then export the workbook in C#. | Show how to keep pivot tables consistent after bulk data updates by using Aspose.Cells' RefreshPivotTables API and writing the result to a new .xlsx file.
// Common Searches: Aspose.Cells C# refresh all pivot tables after updating worksheet data | How to programmatically refresh pivot tables in an Excel file using .NET | Bulk data import then refresh pivot tables with Aspose.Cells example | Using RefreshPivotTables method for multiple worksheets in C#
// Tags: aspocells refreshpivottables method | c# bulk data import excel | excel pivot tables programmatic refresh | aspocells workbook save after pivot refresh | c# update worksheet cells and refresh pivots

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook or creates a new one, fills rows 2‑1000 with sample data, calls workbook.Worksheets.RefreshPivotTables() to update all pivot tables across worksheets, and saves the result to OutputData.xlsx.
public class RefreshAllPivotTablesAfterImport
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string inputPath = "InputData.xlsx";
        const string outputPath = "OutputData.xlsx";

        Workbook workbook;

        // Load existing workbook if it exists; otherwise create a new one.
        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
            workbook.Worksheets[0].Name = "Data";
        }

        // Example of bulk data import – modify many cells as needed
        Worksheet dataSheet = workbook.Worksheets[0];
        for (int row = 2; row <= 1000; row++)
        {
            dataSheet.Cells[$"A{row}"].PutValue($"Item{row}");
            dataSheet.Cells[$"B{row}"].PutValue(row * 10);
        }

        // Refresh every PivotTable in every worksheet to reflect the new data
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook with refreshed pivot tables
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
