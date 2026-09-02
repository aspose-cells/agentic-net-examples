// Title: Log initialized cell count per worksheet and highlight sheets that exceed a specified threshold with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens an Excel file, iterates through each worksheet, counts non‑empty cells, writes the count to the console, and if the count is greater than a given threshold, sets the worksheet tab color to red and fills the used range with a light‑yellow background. | Modify the existing Aspose.Cells program to read the cell‑count threshold from a command‑line argument, allow the highlight color and background style to be configured, and ensure the cell count is logged for every sheet.
// Common Searches: Aspose.Cells count non‑empty cells per worksheet in C# | C# highlight Excel worksheet tab when cell count exceeds limit using Aspose.Cells | apply background style to used range based on cell count Aspose.Cells | log worksheet initialized cell numbers and save modified workbook .NET | set worksheet tab color programmatically with Aspose.Cells after counting cells
// Tags: count initialized cells per worksheet Aspose.Cells | highlight worksheets exceeding cell count threshold | apply tab color and background style Aspose.Cells | log worksheet cell counts .NET | use MaxDisplayRange to style used range Aspose.Cells | configure cell count threshold command line C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, counts non‑empty cells in each worksheet, logs the counts, and for any sheet whose initialized cell count exceeds a defined threshold, it colors the tab red and applies a light‑yellow background to the used range before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the threshold for initialized cells
            int threshold = 1000;

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Count initialized (non‑empty) cells
                int initializedCount = 0;
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        initializedCount++;
                    }
                }

                // Log the count
                Console.WriteLine($"Worksheet '{sheet.Name}' has {initializedCount} initialized cells.");

                // Highlight worksheets exceeding the threshold
                if (initializedCount > threshold)
                {
                    // Set the worksheet tab color to red
                    sheet.TabColor = Color.Red;

                    // Create a style for highlighting
                    Style highlightStyle = workbook.CreateStyle();
                    highlightStyle.ForegroundColor = Color.LightYellow;
                    highlightStyle.Pattern = BackgroundType.Solid;

                    StyleFlag flag = new StyleFlag
                    {
                        CellShading = true
                    };

                    try
                    {
                        // Apply the style to the used range
                        Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                        usedRange.ApplyStyle(highlightStyle, flag);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to apply style on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
