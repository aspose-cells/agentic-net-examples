using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Category");
        worksheet.Cells["C1"].PutValue("Price");

        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue("Fruit");
        worksheet.Cells["C2"].PutValue(1.20);

        worksheet.Cells["A3"].PutValue("Carrot");
        worksheet.Cells["B3"].PutValue("Vegetable");
        worksheet.Cells["C3"].PutValue(0.80);

        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue("Fruit");
        worksheet.Cells["C4"].PutValue(1.10);

        // Apply AutoFilter to the header row (row index 0, columns A‑C)
        // SetRange(startRow, startColumn, endColumn)
        worksheet.AutoFilter.SetRange(0, 0, 2);

        // Freeze the header row so filter controls stay visible while scrolling
        // FreezePanes(row, column, freezedRows, freezedColumns)
        // Freeze the first row (row index 1) and no columns
        worksheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook
        workbook.Save("AutoFilterAndFreeze.xlsx");
    }
}