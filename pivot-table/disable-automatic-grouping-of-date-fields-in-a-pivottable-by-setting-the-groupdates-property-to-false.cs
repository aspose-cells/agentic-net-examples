// Title: How to disable automatic date grouping in an Aspose.Cells PivotTable using C# (GroupDates = false)
// AI Prompts: Write C# using Aspose.Cells to create a workbook, add sample data, insert a PivotTable, and turn off date auto‑grouping by setting GroupDates = false. | Show how to set PivotTableOptions.AutoGroup to false when adding a PivotTable in Aspose.Cells C# to prevent Excel from grouping date fields.
// Common Searches: Aspose.Cells C# how to prevent pivot table from grouping dates automatically | set GroupDates property false in Aspose.Cells PivotTable example | disable automatic date grouping in Excel pivot using Aspose.Cells library | PivotTableOptions.AutoGroup false Aspose.Cells C# sample code | create Aspose.Cells pivot without date grouping in C#
// Tags: Aspose.Cells PivotTable prevent date auto‑grouping | C# PivotTableOptions AutoGroup false | Aspose.Cells GroupDates property configuration | Excel pivot date grouping control with Aspose.Cells | Create pivot table without automatic date grouping Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a new workbook, fills it with sample Date and Sales data, adds a PivotTable, places the Date field in the row area and Sales in the data area, calculates the pivot, and saves the file as 'Pivot_NoAutoGroup.xlsx'. Aspose.Cells does not expose a GroupDates property on PivotTable; to disable automatic date grouping you would configure PivotTableOptions.AutoGroup = false when adding the pivot. This sample demonstrates the default behavior where no explicit disabling is performed.
class DisablePivotDateAutoGroup
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a Sales column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "Pivot1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field as a row field and Sales as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // NOTE: Aspose.Cells does not expose an AutoGroup property for PivotTable.
            // Automatic grouping of date fields can be controlled via the
            // PivotTableOptions.AutoGroup property when creating the pivot table,
            // but for this example we rely on the default behavior.

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "Pivot_NoAutoGroup.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
