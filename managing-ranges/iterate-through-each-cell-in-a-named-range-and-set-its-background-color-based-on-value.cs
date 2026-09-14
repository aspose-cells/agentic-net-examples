// Title: Color cells in a named range based on numeric values using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loops through each cell in a named range and applies LightGreen for values > 0, LightSalmon for values < 0, and LightGray for non‑numeric cells. | Update the provided Aspose.Cells sample to also give text cells a LightBlue background while keeping the existing numeric coloring logic. | Create a reusable C# method that receives a workbook and a range name, then sets conditional background colors for cells based on their type and numeric sign using Aspose.Cells.
// Common Searches: Aspose.Cells C# iterate named range and apply fill color based on cell value | How to set conditional background colors for cells in a specific range with Aspose.Cells for .NET | C# Aspose.Cells change cell background for positive and negative numbers in a named range | Retrieve a named range by name and format its cells using Aspose.Cells API | Apply solid background pattern to Excel cells with Aspose.Cells in C#
// Tags: conditional background coloring Aspose.Cells C# | named range iteration Aspose.Cells | set cell style based on numeric value Aspose.Cells | GetRangeByName Aspose.Cells example | apply solid fill pattern Aspose.Cells | Excel workbook cell formatting .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', gets the named range 'MyRange', iterates each cell, sets LightGreen for positive numbers, LightSalmon for negative numbers, LightGray for non‑numeric cells, and saves the result as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range by its defined name
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("MyRange");
            if (range == null)
            {
                Console.WriteLine("Named range 'MyRange' was not found in the workbook.");
                return;
            }

            // Iterate through each cell in the named range
            foreach (Cell cell in range)
            {
                // Apply background color based on cell value type
                Style style = cell.GetStyle();

                if (cell.Type == CellValueType.IsNumeric)
                {
                    double value = cell.DoubleValue;
                    style.ForegroundColor = value > 0 ? Color.LightGreen : Color.LightSalmon;
                }
                else
                {
                    // Non‑numeric cells get a light gray background
                    style.ForegroundColor = Color.LightGray;
                }

                style.Pattern = BackgroundType.Solid;
                cell.SetStyle(style);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
