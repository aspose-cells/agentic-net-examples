using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ShowHideSubtotalsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ------------------- Regular Subtotals -------------------
        // Populate sample data
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            cells["A" + i].PutValue(i % 2 == 0 ? "Group1" : "Group2");
            cells["B" + i].PutValue(i * 100);
        }

        // Define the range for subtotals (A1:B6)
        CellArea area = CellArea.CreateCellArea("A1", "B6");

        // Add subtotals with summary below data (visible)
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);
        SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
        Console.WriteLine("Regular subtotal - SummaryBelowData (visible): " + setting.SummaryBelowData);

        // Add subtotals again but hide the summary (summaryBelowData = false)
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, false);
        setting = cells.RetrieveSubtotalSetting(area);
        Console.WriteLine("Regular subtotal - SummaryBelowData (hidden): " + setting.SummaryBelowData);

        // ------------------- Pivot Table Subtotals -------------------
        // Populate data for pivot table
        cells["D1"].PutValue("Product");
        cells["E1"].PutValue("Region");
        cells["F1"].PutValue("Sales");
        string[] products = { "Bike", "Bike", "Car", "Car" };
        string[] regions = { "North", "South", "North", "South" };
        int[] sales = { 1000, 1500, 2000, 2500 };
        for (int i = 0; i < products.Length; i++)
        {
            cells[$"D{2 + i}"].PutValue(products[i]);
            cells[$"E{2 + i}"].PutValue(regions[i]);
            cells[$"F{2 + i}"].PutValue(sales[i]);
        }

        // Create a pivot table based on the data range D1:F5
        int pivotIdx = sheet.PivotTables.Add("D1:F5", "H3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Add row field (Product) and data field (Sales)
        int rowFieldIdx = pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        PivotField rowField = pivot.RowFields[rowFieldIdx];
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Show automatic subtotals (default behavior)
        rowField.IsAutoSubtotals = true;
        Console.WriteLine("Pivot subtotal initially visible: " + rowField.IsAutoSubtotals);

        // Hide subtotals for the row field
        rowField.IsAutoSubtotals = false;
        Console.WriteLine("Pivot subtotal after hiding: " + rowField.IsAutoSubtotals);

        // Refresh and calculate the pivot table to apply changes
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("ShowHideSubtotalsDemo.xlsx");
    }
}