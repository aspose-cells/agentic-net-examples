// Title: How to apply a Data Bar conditional format with the workbook’s Accent4 theme color for positive values using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an Excel workbook, fills column A with numbers, and adds a DataBar conditional formatting rule that uses the workbook’s fourth accent color for positive bars using Aspose.Cells. | Write a .NET snippet that applies a data‑bar conditional format to range A1:A10, sets the bar color to ThemeColorType.Accent4, hides negative bars, and saves the workbook.
// Common Searches: Aspose.Cells .NET set data bar color to theme accent4 for positive values | How to hide negative data bars and use Accent4 color in Excel with Aspose.Cells | Apply conditional formatting data bar using GetThemeColor in C# | Create data bar conditional format for range A1:A10 in Aspose.Cells | Use ThemeColorType.Accent4 for data bar styling in Aspose.Cells workbook
// Tags: apply data bar conditional format aspocells | accent4 theme color data bar aspocells | positive values only data bar aspocells | hide negative bars aspocells .net | GetThemeColor usage aspocells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, writes values -5 to 4 into cells A1:A10, defines that range, adds a DataBar conditional formatting rule, sets its color to the workbook’s Accent4 theme color, optionally hides negative bars, displays the cell value beside the bar, and saves the file as DataBarAccent4.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            var workbook = new Workbook();

            // Get the first worksheet.
            var sheet = workbook.Worksheets[0];

            // Populate sample data in column A (some negative, some positive).
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i - 5); // Values -5 to 4.
            }

            // Define the range that will receive the data‑bar conditional format.
            var range = sheet.Cells.CreateRange("A1:A10");

            // Add a new ConditionalFormatting object to the worksheet.
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];

            // Associate the range with the conditional formatting (requires CellArea).
            var area = new CellArea
            {
                StartRow = range.FirstRow,
                StartColumn = range.FirstColumn,
                EndRow = range.FirstRow + range.RowCount - 1,
                EndColumn = range.FirstColumn + range.ColumnCount - 1
            };
            cf.AddArea(area);

            // Add a DataBar condition.
            int conditionIndex = cf.AddCondition(FormatConditionType.DataBar);
            var condition = cf[conditionIndex];
            var dataBar = condition.DataBar;

            // Use the theme's Accent4 color for the positive data bar.
            dataBar.Color = workbook.GetThemeColor(ThemeColorType.Accent4);

            // Hide negative bars (optional – makes the bar appear only for positive values).
            // If the API version supports NegativeBarColor, uncomment the line below.
            // dataBar.NegativeBarColor = Color.Transparent;

            // Show the cell value next to the bar.
            dataBar.ShowValue = true;

            // Save the workbook.
            string outputPath = "DataBarAccent4.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
