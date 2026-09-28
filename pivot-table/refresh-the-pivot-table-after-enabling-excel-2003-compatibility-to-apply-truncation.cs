// Title: C# example: Refresh Aspose.Cells pivot table with Excel 2003 compatibility to truncate strings over 255 characters
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, enables IsExcel2003Compatible, and then calls RefreshData and CalculateData to apply 255‑character truncation using Aspose.Cells. | Show how to set Excel 2003 compatibility on an Aspose.Cells pivot table and refresh its cache so that long text fields are automatically shortened.
// Common Searches: how to truncate text longer than 255 characters in an Aspose.Cells pivot table C# | Aspose.Cells set IsExcel2003Compatible and refresh pivot cache | C# refresh pivot table after enabling Excel 2003 compatibility mode | Aspose.Cells pivot table 2003 compatibility example | RefreshData CalculateData methods for Aspose.Cells pivot tables
// Tags: Aspose.Cells pivot table refresh | Excel 2003 compatibility pivot truncation | IsExcel2003Compatible property C# | RefreshData method Aspose.Cells | CalculateData method pivot table | truncate long strings Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook with source data, adds a pivot table on a separate sheet, enables Excel 2003 compatibility (which truncates strings longer than 255 characters), and then calls RefreshData() and CalculateData() to apply the truncation before saving the file.
public class RefreshPivotTableExcel2003CompatibilityDemo
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
        // Create a new workbook and get the first worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate source data; include a description longer than 255 characters
        string longDescription = new string('X', 300); // 300 characters
        dataSheet.Cells["A1"].Value = "Product";
        dataSheet.Cells["B1"].Value = "Description";
        dataSheet.Cells["A2"].Value = "Item1";
        dataSheet.Cells["B2"].Value = longDescription;
        dataSheet.Cells["A3"].Value = "Item2";
        dataSheet.Cells["B3"].Value = "Short description";

        // Add a new worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table using the data range from the "Data" sheet
        // Parameters: source range, destination cell, pivot table name
        int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B3", "A4", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot table: Product as row field, Description as data field (count)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Description

        // Enable Excel 2003 compatibility mode (truncates strings >255 chars)
        pivotTable.IsExcel2003Compatible = true;

        try
        {
            // Refresh the pivot cache and recalculate the pivot table to apply truncation
            pivotTable.RefreshData();      // Correct API call
            pivotTable.CalculateData();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Pivot refresh error: {ex.Message}");
        }

        // Save the workbook with the refreshed pivot table
        string outputPath = "PivotTable_Excel2003Compatibility.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
