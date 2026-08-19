// Title: Map a collection to a worksheet range using Aspose.Cells Range.Value in C#
// Description: Demonstrates how to convert a List<object[]> into a two‑dimensional array, create a matching Range with Cells.CreateRange, assign the array via Range.Value, and save the workbook. Ideal for populating a defined cell block in a new worksheet.
// Keywords: Aspose.Cells C# Range.Value | CreateRange example | populate Excel range from collection | write 2D object array to worksheet | smart markers Aspose.Cells | map list to Excel cells
// Common Searches: Aspose.Cells set Range.Value from 2D array C# | CreateRange method usage Aspose.Cells | populate Excel sheet with List<object[]> | write collection to specific cell block Aspose | C# map data to Excel range Aspose.Cells
// Developer Intent: Fill a specific block of cells in an Excel worksheet with data from a collection using Aspose.Cells.
// Use Cases: Export header and rows from a List<object[]> to the A1‑based range of a new workbook. | Dynamically size the target range based on the collection’s row and column count. | Save the populated workbook after mapping the data to the worksheet.
// AI Prompts: Generate C# code that converts a List<object[]> to an object[,] and assigns it to an Aspose.Cells Range using Range.Value. | Show how to create a range starting at cell B2 that matches a data collection’s dimensions and populate it with Aspose.Cells. | Explain exception handling best practices when mapping a collection to a worksheet range with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to convert a List<object[]> into a two‑dimensional array, create a matching Range with Cells.CreateRange, assign the array via Range.Value, and save the workbook. Ideal for populating a defined cell block in a new worksheet.
class Program
{
    static void Main()
    {
        try
        {
            // Sample collection of objects (each inner array represents a row)
            var data = new List<object[]>
            {
                new object[] { "Id", "Name", "Score" },
                new object[] { 1, "Alice", 85 },
                new object[] { 2, "Bob", 92 }
            };

            // Convert the collection to a 2‑dimensional array required by Range.Value
            int rowCount = data.Count;
            int colCount = data[0].Length;
            object[,] values = new object[rowCount, colCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    values[i, j] = data[i][j];
                }
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range that matches the size of the data (starting at A1)
            AsposeRange range = cells.CreateRange(0, 0, rowCount, colCount); // rows, columns

            // Map the collection to the worksheet by setting the range's value
            range.Value = values;

            // Save the workbook
            string outputPath = "MappedData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
