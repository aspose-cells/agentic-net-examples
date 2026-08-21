// Title: Aspose.Cells C# – Show Row Grand Totals in a PivotTable
// Description: This example creates a workbook, fills it with sample sales data, adds a PivotTable on range A1:C5, assigns Category, Region, and Sales to row, column, and data fields, enables row grand totals with the ShowRowGrandTotals property, and saves the file as PivotTableShowGrandTotalsForRows.xlsx.
// Keywords: Aspose.Cells PivotTable row totals | C# ShowRowGrandTotals | Enable row grand totals Aspose | PivotTable ShowRowGrandTotals property | Aspose.Cells example C# | Excel pivot table grand totals code | Aspose.Cells reporting
// Common Searches: Aspose.Cells how to display row grand totals | C# set ShowRowGrandTotals in PivotTable | Enable row totals in Aspose.Cells pivot table | PivotTable row grand total property C# | Aspose.Cells sample for row grand totals
// Developer Intent: Add a PivotTable and turn on row grand totals using Aspose.Cells in C#.
// Use Cases: Generate a sales summary that lists each category with a total row. | Create financial expense reports where each group shows a row subtotal. | Export Excel dashboards that require row‑level grand totals for downstream analysis.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable and enable row grand totals. | Explain how the ShowRowGrandTotals property changes the layout of an Aspose.Cells PivotTable. | Provide a complete Aspose.Cells example that sets ShowRowGrandTotals to true and customizes row, column, and data fields.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, fills it with sample sales data, adds a PivotTable on range A1:C5, assigns Category, Region, and Sales to row, column, and data fields, enables row grand totals with the ShowRowGrandTotals property, and saves the file as PivotTableShowGrandTotalsForRows.xlsx.
class SetPivotTableShowGrandTotalsForRows
{
    public static void Main()
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
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        var cells = sheet.Cells;
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Region";
        cells["C1"].Value = "Sales";

        cells["A2"].Value = "Electronics";
        cells["B2"].Value = "North";
        cells["C2"].Value = 1200;

        cells["A3"].Value = "Electronics";
        cells["B3"].Value = "South";
        cells["C3"].Value = 1500;

        cells["A4"].Value = "Furniture";
        cells["B4"].Value = "North";
        cells["C4"].Value = 800;

        cells["A5"].Value = "Furniture";
        cells["B5"].Value = "South";
        cells["C5"].Value = 950;

        // Add a pivot table based on the data range
        int ptIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

        // Enable grand totals for rows
        pivotTable.ShowRowGrandTotals = true;

        // Save the workbook to a file
        string outputPath = "PivotTableShowGrandTotalsForRows.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
