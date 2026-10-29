// Title: Enable PivotField AutoSort to sort row items by a related data field using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable AutoSort on a PivotField to order categories by sales | programmatic way to sort pivot table row items by a related value field in .NET | C# code example for configuring PivotField.AutoSort property with Aspose.Cells | how to make pivot rows automatically sorted by summed data using Aspose.Cells API | sorting Excel pivot rows by a data field through Aspose.Cells library
// Tags: Aspose.Cells PivotField AutoSort C# | pivot table row autosort Aspose.Cells | configure PivotField AutoSort .NET | sort pivot rows by data field Aspose | Excel pivot autosort example C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example builds a workbook, populates it with Category, Product, and Sales data, creates a pivot table with 'Category' as a row field and a summed 'Sales' data field, and shows how to set the PivotField.AutoSort property so that the row items are ordered based on the sales values before saving the file as PivotWithAutoSort.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(150);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Broccoli");
            sheet.Cells["C5"].PutValue(90);

            // Add a pivot table (source range A1:C5, destination starts at A7).
            int pivotIndex = sheet.PivotTables.Add("PivotTable1", "A1:C5", "A7");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add "Category" as a row field (first column in source range).
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);

            // Add "Sales" as a data field (third column in source range) with default sum aggregation.
            pivotTable.DataFields.Add(pivotTable.DataFields[2]);
            // Optionally set a custom name for the data field.
            pivotTable.DataFields[pivotTable.DataFields.Count - 1].Name = "Sum of Sales";

            // Save the workbook.
            workbook.Save("PivotWithAutoSort.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
