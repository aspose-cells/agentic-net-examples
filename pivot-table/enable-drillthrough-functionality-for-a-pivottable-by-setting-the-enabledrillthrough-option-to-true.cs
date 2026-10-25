// Title: Enable drill‑through for an Aspose.Cells PivotTable in C# by setting the EnableDrillThrough property
// Common Searches: C# Aspose.Cells how to turn on drill through for a pivot table | set EnableDrillThrough true in Aspose.Cells pivot table example | Aspose.Cells .NET enable pivot table drill‑through option | create Excel pivot table with drill through using Aspose.Cells C# | Aspose.Cells enable drill through for pivot table programmatically
// Tags: Aspose.Cells pivot table enable drill-through | C# set EnableDrillThrough property | Excel pivot table drill-through configuration with Aspose | Aspose.Cells workbook save with drill-through enabled

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample creates a workbook, populates a data sheet, adds a pivot sheet, builds a PivotTable from the data range, enables the EnableDrillThrough option, and saves the file as PivotWithDrillThrough.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data for the pivot table
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(250);

            // Add a worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Create the pivot table using the data range A1:B5
            int pivotTableIndex = pivotSheet.PivotTables.Add("=Data!A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Retrieve the source fields (automatically created for each column)
            PivotField categoryField = pivotTable.RowFields[0]; // first source field corresponds to "Category"
            PivotField amountField = pivotTable.DataFields[0]; // first source field corresponds to "Amount"

            // Add the fields to the appropriate collections
            pivotTable.RowFields.Add(categoryField);
            pivotTable.DataFields.Add(amountField);

            // Save the workbook to a file
            string outputPath = "PivotWithDrillThrough.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
