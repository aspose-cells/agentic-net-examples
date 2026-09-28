// Title: Add a green data‑bar conditional format to column L for progress percentages using Aspose.Cells for .NET (C#)
// AI Prompts: Create a DataBar conditional formatting rule for column L, set the bar color to green, and display the cell value using Aspose.Cells in C#. | Apply a 0‑100 % length data bar to the range L1:L{lastRow} and configure its appearance with the Aspose.Cells API. | Programmatically add a DataBar format to column L of an existing workbook and save the modified file.
// Common Searches: how to add a green data bar to column L in Excel with Aspose.Cells C# | Aspose.Cells conditional formatting data bar for progress percentage column | C# code to apply data bar conditional formatting to a specific column using Aspose.Cells | set data bar min and max length 0 100 in Aspose.Cells workbook | apply conditional formatting to column L range in existing Excel file Aspose.Cells
// Tags: Aspose.Cells DataBar API | Excel column L conditional formatting | green progress bar in workbook | set data bar min max length Aspose | conditional formatting range definition Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, determines the last used row, defines a range covering column L, adds a DataBar conditional formatting rule, sets the bar color to green, shows the cell value, configures the bar length from 0 % to 100 %, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths (replace placeholders with actual paths)
            string inputFilePath = "{InputFilePath}";
            string outputFilePath = "{OutputFilePath}";

            try
            {
                // Verify that the input workbook exists
                if (string.IsNullOrWhiteSpace(inputFilePath) || !File.Exists(inputFilePath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputFilePath}");
                }

                // Load the existing workbook
                var workbook = new Workbook(inputFilePath);

                // Get the first worksheet (adjust index if needed)
                var sheet = workbook.Worksheets[0];

                // Determine the last used row in the sheet
                int lastRow = sheet.Cells.MaxDataRow;

                // Define the range that covers column L (zero‑based index 11) from the first to the last used row
                var range = new CellArea
                {
                    StartRow = 0,
                    EndRow = lastRow,
                    StartColumn = 11,   // Column L
                    EndColumn = 11
                };

                // Add a new Conditional Formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                var cf = sheet.ConditionalFormattings[cfIndex];

                // Associate the defined range with the conditional formatting
                cf.AddArea(range);

                // Add a Data Bar condition (no operator needed for DataBar)
                int conditionIndex = cf.AddCondition(
                    FormatConditionType.DataBar,
                    OperatorType.None,
                    null,
                    null);

                // Retrieve the condition object
                var condition = cf[conditionIndex];

                // Configure the Data Bar appearance
                var dataBar = condition.DataBar;
                dataBar.Color = Color.FromArgb(0, 176, 80); // Green bar
                dataBar.ShowValue = true;          // Show the cell value next to the bar
                dataBar.MinLength = 0;             // Minimum length percentage
                dataBar.MaxLength = 100;           // Maximum length percentage

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
