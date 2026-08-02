using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Value";
        cells["A2"].Value = "A";
        cells["B2"].Value = 10;
        cells["A3"].Value = "B";
        cells["B3"].Value = 30;
        cells["A4"].Value = "C";
        cells["B4"].Value = 20;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add a row field and a data field to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");

        // Retrieve the data field and set its display format to rank smallest-to-largest (ascending ranking)
        PivotField dataField = pivot.DataFields[0];
        dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankSmallestToLargest;

        // Refresh the pivot table data and calculate the results
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotRankSmallestToLargest.xlsx");
    }
}