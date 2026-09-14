// Title: Create a pivot table with automatic default styling using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a pivot table from a data range and activates the IsAutoFormat property with Aspose.Cells. | Show how to calculate pivot table data after enabling automatic visual styling in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# enable auto format on newly created pivot table | Set default pivot table style programmatically with Aspose.Cells .NET | How to turn on IsAutoFormat for a pivot table using Aspose.Cells API | Create pivot table and apply built‑in visual style automatically in C#
// Tags: Aspose.Cells pivot table auto styling | C# enable pivot auto formatting Aspose.Cells | Aspose.Cells default pivot visual style | programmatic pivot table formatting .NET | calculate pivot data Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds sample data, builds a pivot table, enables automatic default styling via the IsAutoFormat flag, calculates the pivot data, and saves the file as PivotTableAutoFormatDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Enable automatic formatting (default visual styling)
        pivotTable.IsAutoFormat = true;

        // Calculate the pivot table data
        pivotTable.CalculateData();

        // Save the workbook with the formatted pivot table
        workbook.Save("PivotTableAutoFormatDemo.xlsx");
    }
}
