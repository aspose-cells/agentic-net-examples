// Title: How to use Aspose.Cells for .NET to apply conditional formatting that highlights cells in a column exceeding a numeric threshold (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, fills column B with numeric data, defines a threshold value, and applies a solid light‑salmon style to any cell whose value is greater than the threshold. | Write a C# program using Aspose.Cells to iterate through a specific worksheet column, evaluate each cell's numeric value against a given limit, and set a custom background color when the condition is met. | Produce an example that saves an .xlsx file where Aspose.Cells applies a style to automatically highlight values above 100 in column B with a solid color.
// Common Searches: Aspose.Cells C# highlight cells greater than a specific number in a column | conditional formatting column based on numeric threshold using Aspose.Cells for .NET | C# Aspose.Cells apply style to cells where value exceeds 100 | programmatically set background color for cells over a limit in Excel with Aspose.Cells | how to iterate through worksheet cells and apply conditional formatting in Aspose.Cells C#
// Tags: Aspose.Cells value‑based cell styling | C# highlight Excel cells over threshold | apply background color to cells exceeding limit Aspose | generate .xlsx with value‑driven cell style | programmatic Excel cell formatting using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// // Creates a workbook, populates column B with multiples of 20, defines a threshold of 100, builds a light‑salmon style, iterates the first ten rows, applies the style to cells with values > threshold, and saves the file as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column B (index 1)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 1].PutValue(i * 20); // 0, 20, 40, ... 180
            }

            // Define the numeric threshold
            double threshold = 100;

            // Create the style to apply when the condition is met (light salmon background)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.LightSalmon;
            highlightStyle.Pattern = BackgroundType.Solid;

            // Apply the style manually to cells that meet the condition
            for (int row = 0; row < 10; row++)
            {
                Cell cell = sheet.Cells[row, 1];
                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double cellValue))
                {
                    if (cellValue > threshold)
                    {
                        cell.SetStyle(highlightStyle);
                    }
                }
            }

            // Save the workbook
            string outputPath = "ConditionalFormatting.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
