// Title: Collapse all row items in an Aspose.Cells PivotTable (C#)
// Description: Creates a workbook, adds a PivotTable on sample data, places the "Category" field in the row area, refreshes the table, then iterates each PivotItem of the row field and sets its IsDetailHidden property to true, producing a collapsed view saved as CollapsedPivotItems.xlsx.
// Keywords: Aspose.Cells | PivotTable | C# | .NET | Collapse rows | IsDetailHidden | PivotItem | hide pivot details | Excel automation | programmatic pivot collapse
// Common Searches: collapse all rows in Aspose.Cells pivot table | set IsDetailHidden for PivotItem C# | hide row field details programmatically Aspose.Cells | iterate PivotItems to collapse pivot rows | Aspose.Cells pivot table collapse example
// Developer Intent: Programmatically hide the detail rows of a PivotTable by setting each row field's PivotItem.IsDetailHidden to true.
// Use Cases: Generate summary reports where category rows start collapsed, showing only totals. | Prepare Excel files for distribution with pivot details hidden to simplify the user view. | Create workbooks that open with row items collapsed, allowing users to expand them as needed.
// AI Prompts: Show C# code to collapse specific PivotItems by caption using Aspose.Cells. | Give an example of toggling IsDetailHidden for a single PivotItem in a PivotTable. | Explain the difference between PivotItem.IsDetailHidden and PivotItem.IsExpanded in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds a PivotTable on sample data, places the "Category" field in the row area, refreshes the table, then iterates each PivotItem of the row field and sets its IsDetailHidden property to true, producing a collapsed view saved as CollapsedPivotItems.xlsx.
class CollapsePivotItemsDemo
{
    static void Main()
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
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Fruit");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue("Vegetable");
        sheet.Cells["B4"].PutValue(200);
        sheet.Cells["A5"].PutValue("Vegetable");
        sheet.Cells["B5"].PutValue(250);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Category as row field, Amount as data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh and calculate the pivot table to populate it
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Collapse each row field item by hiding its detail
        PivotField rowField = pivotTable.RowFields[0];
        foreach (PivotItem item in rowField.PivotItems)
        {
            item.IsDetailHidden = true;
        }

        // Save the workbook with the collapsed pivot items
        string outputPath = "CollapsedPivotItems.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
