// Title: Set ShowValuesSetting.CalculationType to RankLargestToSmallest for All Pivot Data Fields (Aspose.Cells C#)
// Description: Creates a workbook, adds sample Category/SubCategory/Amount data, builds a pivot table, adds two Amount data fields, loops through the pivot table's DataFields collection, and sets each field's ShowValuesSetting.CalculationType to RankLargestToSmallest before refreshing, recalculating, and saving the file.
// Keywords: Aspose.Cells C# pivot table | ShowValuesSetting CalculationType | RankLargestToSmallest | set calculation type for multiple data fields | loop through pivot data fields | refresh pivot table Aspose.Cells | Excel ranking pivot data | Aspose.Cells API ShowValuesSetting
// Common Searches: Aspose.Cells set ShowValuesSetting to RankLargestToSmallest | C# pivot table rank values largest to smallest | apply calculation type to all data fields Aspose.Cells | loop over PivotTable.DataFields C# | change pivot field display format Aspose.Cells
// Developer Intent: Apply the RankLargestToSmallest calculation type uniformly to every data field in an Aspose.Cells pivot table using C#.
// Use Cases: Rank multiple data fields in a pivot table without configuring each field separately. | Update the display calculation of existing pivot data fields after they have been added. | Refresh and recalculate a pivot table to reflect a new ranking order.
// AI Prompts: Generate C# code with Aspose.Cells that sets ShowValuesSetting.CalculationType to RankLargestToSmallest for all pivot data fields. | Explain how ShowValuesSetting.CalculationType affects pivot table output in Aspose.Cells and demonstrate an efficient way to apply it to several fields. | Provide a LINQ one‑liner to assign RankLargestToSmallest to each PivotField in pivotTable.DataFields.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds sample Category/SubCategory/Amount data, builds a pivot table, adds two Amount data fields, loops through the pivot table's DataFields collection, and sets each field's ShowValuesSetting.CalculationType to RankLargestToSmallest before refreshing, recalculating, and saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "SubCategory";
        sheet.Cells["C1"].Value = "Amount";

        string[] categories = { "A", "A", "B", "B", "C", "C" };
        string[] subCategories = { "X", "Y", "X", "Y", "X", "Y" };
        double[] amounts = { 100, 200, 150, 250, 120, 220 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 1, 0].Value = categories[i];
            sheet.Cells[i + 1, 1].Value = subCategories[i];
            sheet.Cells[i + 1, 2].Value = amounts[i];
        }

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
        // Add two data fields (same source column) to demonstrate uniform setting
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Set ShowValuesSetting.CalculationType to RankLargestToSmallest for all data fields
        foreach (PivotField dataField in pivotTable.DataFields)
        {
            dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;
        }

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotShowValuesRankLargestToSmallest.xlsx");
    }
}
