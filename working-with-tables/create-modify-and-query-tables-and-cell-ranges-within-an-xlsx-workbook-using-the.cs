using System;
using System.Data;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the default worksheet and give it a friendly name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Demo";

        // Populate a simple product table (A1:C4)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(0.5);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue(0.3);

        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(15);
        sheet.Cells["C4"].PutValue(0.8);

        // Create a range that covers the data table
        Aspose.Cells.Range dataRange = sheet.Cells.CreateRange("A1", "C4");

        // Apply a header style to the first row of the range
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        StyleFlag headerFlag = new StyleFlag
        {
            FontBold = true,
            CellShading = true
        };
        dataRange.ApplyStyle(headerStyle, headerFlag);

        // Insert a new column (D) for Total = Quantity * Price
        sheet.Cells.InsertColumn(3); // zero‑based index, column D
        sheet.Cells["D1"].PutValue("Total");

        // Set formula for each data row in the Total column
        for (int row = 2; row <= 4; row++)
        {
            // Row and column indexes are zero‑based for the Cells indexer
            sheet.Cells[row - 1, 3].Formula = $"=B{row}*C{row}";
        }

        // Calculate formulas so the Total column is populated
        workbook.CalculateFormula();

        // Clear the contents of the second data row (row 3) without removing formatting
        // Parameters: startRow, startColumn, endRow, endColumn (zero‑based)
        sheet.Cells.ClearContents(2, 0, 2, 2);

        // Merge cells A6:C6 to create a title area
        sheet.Cells.Merge(5, 0, 1, 3);
        sheet.Cells["A6"].PutValue("Sales Summary");

        // Apply a centered, bold style to the merged title cell
        Style titleStyle = workbook.CreateStyle();
        titleStyle.Font.IsBold = true;
        titleStyle.Font.Size = 14;
        titleStyle.HorizontalAlignment = TextAlignmentType.Center;
        sheet.Cells["A6"].SetStyle(titleStyle);

        // Query specific cell values
        string product = sheet.Cells["A2"].StringValue; // Expected "Apple"
        double total = sheet.Cells["D2"].DoubleValue;   // Expected 10 * 0.5 = 5

        Console.WriteLine($"Product: {product}, Total: {total}");

        // Export the full table (including the Total column) to a DataTable
        // Parameters: startRow, startColumn, totalRows, totalColumns, exportColumnNames
        DataTable dt = sheet.Cells.ExportDataTable(0, 0, 5, 4, true);
        Console.WriteLine("Exported DataTable rows:");
        foreach (DataRow r in dt.Rows)
        {
            Console.WriteLine(string.Join(", ", r.ItemArray));
        }

        // Save the workbook to disk (lifecycle rule)
        workbook.Save("DemoWorkbook.xlsx");
    }
}