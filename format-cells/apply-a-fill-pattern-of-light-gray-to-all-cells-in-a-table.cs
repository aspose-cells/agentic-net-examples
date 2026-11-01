// Title: How to apply a light gray solid fill pattern to an entire Excel table range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set solid light gray background for a specific range of cells | Apply fill pattern to all cells in an Excel table with Aspose.Cells .NET | How to use StyleFlag to format a table range in Aspose.Cells C# | Change cell background color to light gray for A1:D10 using Aspose.Cells | Aspose.Cells create and style a range with solid fill in C#
// Tags: apply solid fill Aspose.Cells C# | style entire range Excel Aspose.Cells | light gray cell background Aspose.Cells | create range A1:D10 Aspose.Cells | use StyleFlag All property Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, defines the range A1:D10, creates a style with a solid light gray foreground fill, applies the style to every cell in the range using a StyleFlag with All=true, and saves the modified workbook to a new file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook. If the file is password‑protected, provide the password here.
                var loadOptions = new LoadOptions(); // Add Password = "yourPassword" if needed.
                var workbook = new Workbook(inputPath, loadOptions);

                // Get the first worksheet (adjust index if needed).
                var worksheet = workbook.Worksheets[0];

                // Define the range that represents the table.
                // Replace "A1:D10" with the actual address of your table.
                var range = worksheet.Cells.CreateRange("A1:D10");

                // Create a style with a light gray solid fill pattern.
                var style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;          // Solid fill.
                style.ForegroundColor = Color.LightGray;      // Light gray color.
                style.BackgroundColor = Color.White;          // Optional background color.

                // Apply the style to all cells in the range.
                var styleFlag = new StyleFlag { All = true };
                range.ApplyStyle(style, styleFlag);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
