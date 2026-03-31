using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the data to be populated (3 rows x 3 columns)
        object[,] data = new object[,]
        {
            { "Product", "Price", "Quantity" },
            { "Apple",   1.2,    10 },
            { "Banana",  0.8,    20 }
        };

        // Populate the range starting at cell A1
        for (int row = 0; row < data.GetLength(0); row++)
        {
            for (int col = 0; col < data.GetLength(1); col++)
            {
                cells[row, col].PutValue(data[row, col]);
            }
        }

        // Save the workbook as an XLSX file (lifecycle save)
        workbook.Save("PopulatedData.xlsx", SaveFormat.Xlsx);
    }
}