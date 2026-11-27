// Title: How to apply a Data Bar conditional format with a Theme Accent4 color for positive values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set data bar conditional format with theme accent color for positive numbers | example of applying Excel data bar formatting using theme colors in .NET | C# code to add data bar conditional formatting to a cell range with Aspose.Cells | how to display only positive values in a data bar conditional format with Aspose.Cells | using Aspose.Cells to create a data bar that uses the workbook’s Accent4 theme color
// Tags: Aspose.Cells data bar conditional formatting .NET | theme accent4 color for Excel data bars C# | positive values data bar conditional format Aspose | apply data bar to cell range using Aspose.Cells | Excel conditional formatting data bar API C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, fills cells A1:A10 with values -5 to 4, adds a DataBar conditional formatting rule that uses the workbook’s Accent4 theme color (shown as blue in the code) and displays the cell value, applies the rule to the range A1:A10, and saves the workbook as DataBarConditionalFormat.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10) with values -5 to 4
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i - 5);
            }

            // Define the range for conditional formatting
            string rangeAddress = "A1:A10";

            // Add a new ConditionalFormatting object to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];

            // Add a DataBar condition
            int conditionIndex = cf.AddCondition(FormatConditionType.DataBar);
            var condition = cf[conditionIndex] as FormatCondition;

            // Configure the DataBar
            var dataBar = condition.DataBar;
            dataBar.Color = Color.Blue;      // Bar color
            dataBar.ShowValue = true;        // Show the cell value next to the bar

            // Apply the conditional formatting to the defined range
            var area = CellArea.CreateCellArea("A1", "A10");
            cf.AddArea(area);

            // Determine output file path
            string outputFile = "DataBarConditionalFormat.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

            // Create output directory if needed
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
