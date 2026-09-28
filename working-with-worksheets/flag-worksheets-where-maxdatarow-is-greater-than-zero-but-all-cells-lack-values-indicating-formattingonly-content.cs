// Title: Identify and flag formatting‑only worksheets in an Excel workbook with Aspose.Cells for .NET – add a custom property and apply a LightCoral tab highlight
// AI Prompts: Write C# code using Aspose.Cells that iterates all worksheets, verifies MaxDataRow > 0 while every cell in the used range is empty, then stores a custom property called "FormattingOnly" and sets the worksheet tab color to LightCoral. | Create a reusable C# method with Aspose.Cells that marks worksheets containing only formatting (no cell values) by recording a flag property and changing the tab appearance, then saves the workbook.
// Common Searches: how to detect worksheets that only have formatting and no data using Aspose.Cells C# | Aspose.Cells mark empty data sheets with a custom property | set Excel worksheet tab color based on content with Aspose.Cells .NET | use MaxDataRow to find formatting‑only sheets in C# | flag worksheets without cell values in Aspose.Cells and change tab highlight
// Tags: identify formatting-only sheets Aspose.Cells | store FormattingOnly flag as custom property Aspose.Cells | apply LightCoral tab highlight Aspose.Cells | evaluate used cell range for empty values Aspose.Cells | MaxDataRow based empty sheet detection Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, loops through each worksheet, and uses MaxDataRow/MaxDataColumn to define the used range. If the range reports rows but all cells are null, the sheet is treated as formatting‑only: a custom property "FormattingOnly" is added and the tab color is changed to LightCoral. The modified workbook is then saved.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range size (zero‑based indices)
                    int maxDataRow = sheet.Cells.MaxDataRow + 1;
                    int maxDataColumn = sheet.Cells.MaxDataColumn + 1;

                    // Proceed only if there is at least one reported data row
                    if (maxDataRow > 0 && maxDataColumn > 0)
                    {
                        bool hasValue = false;

                        // Scan cells within the used range
                        for (int row = 0; row < maxDataRow && !hasValue; row++)
                        {
                            for (int col = 0; col < maxDataColumn; col++)
                            {
                                if (sheet.Cells[row, col].Value != null)
                                {
                                    hasValue = true;
                                    break;
                                }
                            }
                        }

                        // If no cell contains a value, treat the sheet as formatting‑only
                        if (!hasValue)
                        {
                            // Store a custom property (value must be a string)
                            sheet.CustomProperties.Add("FormattingOnly", true.ToString());

                            // Change the tab color as a visual indicator
                            sheet.TabColor = Color.LightCoral;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
