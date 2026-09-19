// Title: Set the workbook’s Dark2 theme color as the default border for every table (ListObject) using Aspose.Cells for .NET
// AI Prompts: Extract the Dark2 color from the workbook’s theme and apply it to all four borders of each ListObject range. | Refactor the existing code to replace the hard‑coded black border with the workbook’s Dark2 theme color for all tables across all worksheets. | Create a reusable method that receives a Workbook and automatically styles every table’s borders with the workbook’s Dark2 theme color.
// Common Searches: Aspose.Cells C# set table border color to workbook theme Dark2 | How to use Excel theme colors for ListObject borders with Aspose.Cells .NET | Apply default border style to all tables in an Excel file using Aspose.Cells | Retrieve Dark2 theme color from workbook and style table borders in C# | Change border color of all tables to theme color in Aspose.Cells
// Tags: set table border color using workbook theme Aspose.Cells | apply Dark2 theme color to ListObject borders C# | iterate worksheets and tables to style borders Aspose.Cells | style flag all cells for table borders .NET | theme color extraction for cell styling Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, obtains the Dark2 color from the workbook’s theme, iterates through every worksheet and each ListObject (table), creates a style that sets the left, right, top, and bottom borders to the Dark2 color, applies this style to the entire table range, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                Workbook workbook;
                try
                {
                    // Load the workbook
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Use a default border color (Black)
                Color borderColor = Color.Black;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all tables (ListObjects) in the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // Get the data range that the table occupies
                        Aspose.Cells.Range tableRange = table.DataRange;

                        // Create a style and set border colors
                        Style style = workbook.CreateStyle();
                        style.Borders[BorderType.LeftBorder].Color = borderColor;
                        style.Borders[BorderType.RightBorder].Color = borderColor;
                        style.Borders[BorderType.TopBorder].Color = borderColor;
                        style.Borders[BorderType.BottomBorder].Color = borderColor;

                        // Apply the style to the entire table range
                        StyleFlag flag = new StyleFlag { All = true };
                        tableRange.ApplyStyle(style, flag);
                    }
                }

                try
                {
                    // Save the modified workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
