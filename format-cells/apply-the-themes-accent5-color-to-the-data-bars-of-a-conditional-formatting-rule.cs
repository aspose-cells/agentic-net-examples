// Title: Set a Data Bar conditional formatting rule to use the workbook’s Accent5 theme color with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set data bar conditional format to workbook theme Accent5 color | How to use Excel theme Accent5 for data bar in Aspose.Cells .NET | Apply workbook theme color to conditional formatting data bar using Aspose.Cells | Retrieve Accent5 color from workbook theme for conditional formatting in Aspose.Cells | C# Aspose.Cells change data bar color to theme accent color
// Tags: Aspose.Cells data bar theme color | C# conditional formatting Accent5 | Excel workbook theme color Aspose.Cells | set data bar color from theme .NET | conditional formatting data bar API Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates a Data Bar conditional formatting rule for the range A1:A10 on the first worksheet, assigns the data bar's color to the workbook's Accent5 theme color (with a fallback to blue), and saves the modified file as a new workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A1:A10 for conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a new conditional formatting entry
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area);

            // Add a DataBar condition
            int conditionIndex = cf.AddCondition(FormatConditionType.DataBar);
            var condition = cf[conditionIndex];
            DataBar dataBar = condition.DataBar;

            // Set the data bar color (fallback to a default color)
            Color barColor = Color.Blue;
            dataBar.Color = barColor;

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
