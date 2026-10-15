// Title: Import a two‑dimensional boolean array into an Excel worksheet and apply green/red background colors with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to write a bool[,] into a worksheet and set each cell's fill to green when the value is true and red when false. | Create a reusable C# method that accepts a 2D boolean array, populates an Excel sheet with the values, and applies green or red background colors based on the boolean using Aspose.Cells.
// Common Searches: Aspose.Cells C# write bool[,] to Excel and color cells based on value | how to set cell fill color programmatically for true/false in Aspose.Cells | C# example for applying green red background to boolean data in Excel with Aspose.Cells | populate Excel worksheet from boolean matrix and style cells using Aspose.Cells .NET
// Tags: boolean array import Aspose.Cells | cell background color based on boolean Aspose.Cells | set cell style programmatically Aspose.Cells | green red fill Aspose.Cells | bool[,] to Excel C#

using System;
using Aspose.Cells;
using System.Drawing;

// // Demonstrates creating a workbook, importing a two‑dimensional bool array into the first worksheet, writing each boolean value to a cell, and applying a green background for true or a red background for false before saving as BooleanColors.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Sample boolean array to import
            bool[,] boolData = new bool[,]
            {
                { true, false, true },
                { false, true, false }
            };

            int rows = boolData.GetLength(0);
            int cols = boolData.GetLength(1);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with boolean values
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    Cell cell = sheet.Cells[i, j];
                    cell.PutValue(boolData[i, j]);

                    // Apply background color based on the boolean value
                    Style style = cell.GetStyle();
                    style.ForegroundColor = boolData[i, j] ? Color.Green : Color.Red;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }
            }

            // Save the workbook to a file
            string outputPath = "BooleanColors.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
