// Title: Use Aspose.Cells for .NET to highlight cells with dd/MM/yyyy dates in the J1:J30 range
// AI Prompts: Write C# code that creates a workbook, defines the range J1:J30, checks each cell with a ^\d{2}/\d{2}/\d{4}$ regex, and applies a yellow background style to matching cells using Aspose.Cells. | Generate a routine in C# using Aspose.Cells that iterates over cells J1 through J30, validates the text against a dd/MM/yyyy pattern, and sets a solid yellow fill for cells that match. | Provide a C# example that loads or creates an Excel file, selects column J rows 1‑30, uses a regular expression to detect date strings in dd/MM/yyyy format, and highlights those cells with Aspose.Cells styling.
// Common Searches: aspocells c# highlight cells in column J that match dd/MM/yyyy pattern | how to apply regex date validation to a specific range using Aspose.Cells .NET | C# Aspose.Cells example for searching dates formatted as dd/MM/yyyy in J1:J30 | using Aspose.Cells to color cells with date strings in a given range
// Tags: Aspose.Cells regex date detection | highlight matching cells yellow Aspose.Cells | iterate over range J1:J30 C# | apply style based on pattern Aspose.Cells | C# date format dd/MM/yyyy validation in Excel

using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a new workbook, defines the range J1:J30, uses a regular expression to identify cells containing dates in dd/MM/yyyy format, applies a solid yellow fill to those cells, and saves the file as Output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target range J1:J30 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
                Aspose.Cells.Range targetRange = worksheet.Cells.CreateRange("J1", "J30");

                // Regular expression for dates in dd/MM/yyyy format
                Regex dateRegex = new Regex(@"^\d{2}/\d{2}/\d{4}$");

                // Prepare a style to highlight matching cells (optional)
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;
                StyleFlag styleFlag = new StyleFlag { All = true };

                // Iterate through each cell in the range and apply the regex
                foreach (Cell cell in targetRange)
                {
                    string cellText = cell.StringValue; // Get the cell's text representation
                    if (dateRegex.IsMatch(cellText))
                    {
                        // Cell matches the date pattern; apply the highlight style
                        cell.SetStyle(highlightStyle, styleFlag);
                    }
                }

                // Define output file path
                string outputPath = "Output.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
