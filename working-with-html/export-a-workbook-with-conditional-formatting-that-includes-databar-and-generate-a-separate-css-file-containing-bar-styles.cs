// Title: Export an Excel workbook with DataBar conditional formatting to HTML and generate an external CSS file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet with DataBar conditional formatting to HTML and separate CSS file | How to save an Excel workbook as HTML and get a standalone stylesheet for data bars using Aspose.Cells | Aspose.Cells HtmlSaveOptions generate external CSS for DataBar conditional formatting in .NET | Export Excel DataBar conditional formatting to HTML while creating a CSS file in C# | Separate CSS file for conditional formatting bars when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HTML export external CSS | DataBar conditional formatting HTML conversion | C# HtmlSaveOptions generate stylesheet | Excel to HTML DataBar CSS generation | Conditional formatting CSS output Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a new workbook, fills cells A1:A10 with numbers, applies a DataBar conditional formatting rule, configures its length range, and saves the workbook as HTML using HtmlSaveOptions. The export automatically produces a separate CSS file that contains the styles for the data bars.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10)
            for (int row = 0; row < 10; row++)
            {
                sheet.Cells[row, 0].PutValue(row + 1);
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var conditionalFormatting = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A10 for the DataBar
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };
            conditionalFormatting.AddArea(area);

            // Create a DataBar condition
            int conditionIndex = conditionalFormatting.AddCondition(FormatConditionType.DataBar);
            var dataBarCondition = conditionalFormatting[conditionIndex];

            // Set minimum and maximum lengths (percentage of the cell width)
            dataBarCondition.DataBar.MinLength = 0;   // 0%
            dataBarCondition.DataBar.MaxLength = 100; // 100%

            // Optionally set the color of the data bar (commented out if not supported in current version)
            // dataBarCondition.DataBar.BarColor = System.Drawing.Color.LightGreen;

            // Prepare HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML (which will also create the CSS file if supported)
            workbook.Save("DataBarWorkbook.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
