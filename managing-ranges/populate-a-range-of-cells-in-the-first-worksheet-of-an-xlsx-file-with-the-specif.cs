using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class PopulateRangeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define a 2D array with sample data to populate the range
        object[,] data = new object[,]
        {
            { "Item", "Quantity", "Price" },
            { "Apple", 10, 0.5 },
            { "Banana", 20, 0.3 }
        };

        // Create a range covering cells A1:C3
        AsposeRange range = cells.CreateRange(0, 0, 3, 3);

        // Populate the range with the data array
        for (int row = 0; row < range.RowCount; row++)
        {
            for (int col = 0; col < range.ColumnCount; col++)
            {
                range[row, col].PutValue(data[row, col]);
            }
        }

        // Save the workbook as an XLSX file
        workbook.Save("PopulatedRange.xlsx", SaveFormat.Xlsx);
    }
}