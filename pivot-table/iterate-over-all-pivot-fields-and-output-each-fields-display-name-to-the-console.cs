// Title: How to enumerate all pivot table fields and print their display names with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, adds a pivot table, refreshes it, and loops through the PivotTable.BaseFields collection to write each field's DisplayName to the console. | Provide a complete example that retrieves the list of pivot field names from an Aspose.Cells PivotTable and outputs them via Console.WriteLine after the pivot data is calculated.
// Common Searches: aspnet c# get list of pivot table field names using Aspose.Cells | how to loop through BaseFields of a PivotTable in Aspose.Cells .NET | display pivot field display names in console with Aspose.Cells | sample code for iterating pivot fields in Aspose.Cells C#
// Tags: Aspose.Cells iterate pivot BaseFields C# | pivot table field display name Aspose.Cells | C# console output pivot field names | create and refresh pivot table Aspose.Cells | Aspose.Cells PivotTable enumeration example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program creates a workbook, populates sample data, adds and refreshes a pivot table, then iterates over the PivotTable.BaseFields collection, writing each field's DisplayName to the console before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Product";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Electronics";
        sheet.Cells["B2"].Value = "Laptop";
        sheet.Cells["C2"].Value = 1200;

        sheet.Cells["A3"].Value = "Electronics";
        sheet.Cells["B3"].Value = "Phone";
        sheet.Cells["C3"].Value = 800;

        sheet.Cells["A4"].Value = "Furniture";
        sheet.Cells["B4"].Value = "Chair";
        sheet.Cells["C4"].Value = 150;

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Iterate over all pivot fields (BaseFields) and output each field's display name
        foreach (PivotField field in pivotTable.BaseFields)
        {
            Console.WriteLine(field.DisplayName);
        }

        // Save the workbook
        workbook.Save("PivotFieldsDisplayNameDemo.xlsx");
    }
}
