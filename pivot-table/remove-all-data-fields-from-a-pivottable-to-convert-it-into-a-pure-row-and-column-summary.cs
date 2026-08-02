// Title: Aspose.Cells C# – Remove All Data Fields from a PivotTable to Keep Only Row & Column Labels
// Description: Shows how to build a PivotTable in a new workbook, iterate the DataFields collection in reverse, delete every data field, refresh and recalculate the pivot, and save the file so that only row and column headings remain.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | remove data fields | clear pivot values | row column summary | delete data fields Aspose | Excel automation | Aspose.Cells API example | GitHub Aspose.Cells pivot
// Common Searches: Aspose.Cells remove data fields from PivotTable C# | How to delete all values in a PivotTable using Aspose.Cells | Create PivotTable with only row and column labels in .NET | Aspose.Cells example to clear pivot data fields | C# code to keep pivot layout but hide values
// Developer Intent: Delete every data field in a PivotTable so that the output contains only the row and column headings.
// Use Cases: Produce a template that shows the hierarchical structure of categories without any numeric totals. | Generate documentation or screenshots of a pivot layout while omitting confidential data. | Create a lightweight summary report where only the row/column framework is needed for further processing.
// AI Prompts: Write C# code with Aspose.Cells that removes all data fields from an existing PivotTable and updates the view. | Explain why iterating PivotTable.DataFields in reverse order prevents index errors when deleting fields. | Show how to modify the sample to retain a specific data field (e.g., "Amount") while removing all others.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to build a PivotTable in a new workbook, iterate the DataFields collection in reverse, delete every data field, refresh and recalculate the pivot, and save the file so that only row and column headings remain.
class RemoveDataFieldsFromPivot
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("SubCategory");
        worksheet.Cells["C1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Fruit");
        worksheet.Cells["B2"].PutValue("Apple");
        worksheet.Cells["C2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Fruit");
        worksheet.Cells["B3"].PutValue("Banana");
        worksheet.Cells["C3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Vegetable");
        worksheet.Cells["B4"].PutValue("Carrot");
        worksheet.Cells["C4"].PutValue(15);

        // Add a pivot table covering the data range and place it at E1
        int pivotIndex = worksheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add fields: Category as rows, SubCategory as columns, Amount as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Build the initial pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Remove all data fields to leave only row/column summaries
        for (int i = pivotTable.DataFields.Count - 1; i >= 0; i--)
        {
            string fieldName = pivotTable.DataFields[i].Name;
            pivotTable.RemoveField(PivotFieldType.Data, fieldName);
        }

        // Recalculate after removal to update the view
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the modified pivot table
        workbook.Save("PivotWithoutDataFields.xlsx");
        Console.WriteLine("Workbook saved as PivotWithoutDataFields.xlsx");
    }
}
