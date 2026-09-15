// Title: Apply red and green conditional formatting to the "HighValue" named range in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that formats cells in the named range "HighValue" with a red background when the value exceeds 1000 and a light‑green background when the value is below 500. | Show how to create a reusable C# method that accepts any named range, custom threshold values, and colors to apply multiple value‑based conditional formatting rules with Aspose.Cells. | Provide a step‑by‑step example that retrieves a named range, adds two cell‑value conditions, and saves the workbook after applying the formatting using Aspose.Cells for .NET.
// Common Searches: aspocells c# conditional formatting on a named range based on value thresholds | how to set red background for cells greater than 1000 in an Aspose.Cells workbook | apply light green fill to cells less than 500 within a specific named range using Aspose.Cells .NET | retrieve a named range and add multiple format conditions with Aspose.Cells C# example
// Tags: Aspose.Cells conditional formatting named range | C# apply cell value thresholds formatting | Excel workbook red background high values Aspose | light green fill low values Aspose.Cells | retrieve named range Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, obtains the named range "HighValue", creates a conditional formatting collection on its worksheet, and defines two rules: values > 1000 receive a solid red background and values < 500 receive a solid light‑green background. The workbook is then saved with the new formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Retrieve the named range "HighValue"
            var range = workbook.Worksheets.GetRangeByName("HighValue");
            if (range == null)
            {
                Console.WriteLine("Named range \"HighValue\" does not exist.");
                return;
            }

            // Get the worksheet that contains the named range
            var worksheet = range.Worksheet;

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var cf = worksheet.ConditionalFormattings[cfIndex];

            // Apply the named range area to the conditional formatting using CellArea
            var area = new CellArea
            {
                StartRow = range.FirstRow,
                StartColumn = range.FirstColumn,
                EndRow = range.FirstRow + range.RowCount - 1,
                EndColumn = range.FirstColumn + range.ColumnCount - 1
            };
            cf.AddArea(area);

            // Condition 1: values greater than 1000 → red background
            int condHighIdx = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "1000", string.Empty);
            var conditionHigh = cf[condHighIdx];
            var styleHigh = conditionHigh.Style;
            styleHigh.ForegroundColor = Color.Red;
            styleHigh.Pattern = BackgroundType.Solid;

            // Condition 2: values less than 500 → light green background
            int condLowIdx = cf.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "500", string.Empty);
            var conditionLow = cf[condLowIdx];
            var styleLow = conditionLow.Style;
            styleLow.ForegroundColor = Color.LightGreen;
            styleLow.Pattern = BackgroundType.Solid;

            // Save the workbook with the applied conditional formatting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
