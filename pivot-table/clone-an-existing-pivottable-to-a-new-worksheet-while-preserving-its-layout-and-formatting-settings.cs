using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");
        for (int i = 2; i <= 10; i++)
        {
            dataSheet.Cells[$"A{i}"].PutValue("Cat" + ((i % 3) + 1));
            dataSheet.Cells[$"B{i}"].PutValue("Prod" + i);
            dataSheet.Cells[$"C{i}"].PutValue(i * 100);
        }

        // Create the original PivotTable on its own worksheet
        Worksheet originalPivotSheet = workbook.Worksheets.Add("OriginalPivot");
        int originalIndex = originalPivotSheet.PivotTables.Add("=Data!A1:C10", "E1", "OriginalPivot");
        PivotTable originalPivot = originalPivotSheet.PivotTables[originalIndex];
        originalPivot.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row
        originalPivot.AddFieldToArea(PivotFieldType.Column, 1);   // Product as column
        originalPivot.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data
        originalPivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        originalPivot.CalculateData();

        // Add a new worksheet where the cloned PivotTable will reside
        Worksheet clonedPivotSheet = workbook.Worksheets.Add("ClonedPivot");

        // Clone the PivotTable using the Add method that takes a source PivotTable
        int clonedIndex = clonedPivotSheet.PivotTables.Add(originalPivot, "A1", "ClonedPivot");
        PivotTable clonedPivot = clonedPivotSheet.PivotTables[clonedIndex];

        // Copy the style explicitly (the Add method already copies layout, this ensures formatting)
        clonedPivot.CopyStyle(originalPivot);

        // Refresh and calculate the cloned PivotTable
        clonedPivotSheet.RefreshPivotTables();
        clonedPivot.CalculateData();

        // Save the workbook
        workbook.Save("ClonedPivotDemo.xlsx");
    }
}