using System;
using Aspose.Cells;
using System.Drawing;

class CellsManipulationDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cells collection
        Cells cells = worksheet.Cells;

        // ----- Put values into cells -----
        // Using row/column index (A1 and B1)
        cells[0, 0].PutValue("Product"); // A1
        cells[0, 1].PutValue("Price");   // B1

        // Using cell name (A2, B2, A3, B3)
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(1.25);
        cells["A3"].PutValue("Banana");
        cells["B3"].PutValue(0.75);

        // ----- Retrieve values from cells -----
        // By index
        string header = cells[0, 0].StringValue;      // "Product"
        double applePrice = cells[1, 1].DoubleValue;  // B2

        // By name
        string banana = cells["A3"].StringValue;      // "Banana"
        double bananaPrice = cells["B3"].DoubleValue; // 0.75

        // Display retrieved values
        Console.WriteLine($"{header} - Apple price: {applePrice}");
        Console.WriteLine($"{banana} - Price: {bananaPrice}");

        // ----- Modify a cell value -----
        // Increase Apple price
        cells["B2"].PutValue(1.30);

        // ----- Apply style to header row -----
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        cells["A1"].SetStyle(headerStyle);
        cells["B1"].SetStyle(headerStyle);

        // Save the workbook (lifecycle rule)
        workbook.Save("ManipulatedCells.xlsx");
    }
}