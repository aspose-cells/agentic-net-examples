using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: a category column and a numeric column representing fractions
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Fraction");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(0.1234);   // 12.34%
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(0.5678);   // 56.78%
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(0.9);      // 90.00%

        // Add a pivot table covering the data range and place it starting at D3
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add the Category field to the row area
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");

        // Add the Fraction field to the data area
        int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Fraction");
        PivotField dataField = pivot.DataFields[dataFieldPos];

        // Define a custom number format that shows values as percentages with two decimal places
        dataField.NumberFormat = "0.00%";

        // Refresh the pivot table data and calculate the results
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotTableCustomNumberFormat.xlsx");
    }
}