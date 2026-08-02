// Title: Apply Accent5 Theme Color to Data Bar Conditional Formatting with Aspose.Cells for .NET
// Description: This example creates a workbook, fills cells A1:A5 with numbers, adds a Data Bar conditional format, retrieves the workbook's Accent5 theme color via GetThemeColor, assigns it to the Data Bar, configures automatic minimum and maximum values, shows the cell value, and saves the result as ThemeAccent5DataBar.xlsx.
// Keywords: Aspose.Cells | C# | DataBar | conditional formatting | theme accent color | GetThemeColor | Accent5 | Excel workbook styling | FormatCondition | XLSX export
// Common Searches: Aspose.Cells set data bar color from theme | GetThemeColor Accent5 C# example | apply workbook theme accent to conditional formatting | data bar conditional formatting Aspose.Cells .NET | change data bar color programmatically in Excel
// Developer Intent: Set a Data Bar conditional formatting rule to use the workbook’s Accent5 theme color.
// Use Cases: Generate reports where data bars automatically match the workbook’s Accent5 color for a cohesive visual theme. | Apply the same accent color to multiple data bar rules across worksheets to maintain consistent branding. | Build interactive dashboards that adapt their conditional formatting colors when the workbook theme is changed.
// AI Prompts: Write C# code that retrieves any theme accent (e.g., Accent3) and applies it to a Data Bar conditional format using Aspose.Cells. | Show how to change the data bar color to the workbook's Accent5 theme color and adjust the conditional range to B2:B10. | Explain how to switch a data bar's color between different theme accents based on user input in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeAccentDemo
{
    // This example creates a workbook, fills cells A1:A5 with numbers, adds a Data Bar conditional format, retrieves the workbook's Accent5 theme color via GetThemeColor, assigns it to the Data Bar, configures automatic minimum and maximum values, shows the cell value, and saves the result as ThemeAccent5DataBar.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample numeric data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(30);
                sheet.Cells["A3"].PutValue(50);
                sheet.Cells["A4"].PutValue(70);
                sheet.Cells["A5"].PutValue(90);

                // Define the range that will receive the data bar conditional formatting
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add an empty conditional formatting collection
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
                cfCollection.AddArea(area);

                // Add a DataBar condition
                int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
                FormatCondition condition = cfCollection[conditionIndex];

                // Retrieve the theme's Accent5 color and apply it to the data bar
                Color accent5 = workbook.GetThemeColor(ThemeColorType.Accent5);
                condition.DataBar.Color = accent5;

                // Set other required properties for the data bar
                condition.DataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
                condition.DataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
                condition.DataBar.ShowValue = true; // display the cell value alongside the bar

                // Save the workbook
                workbook.Save("ThemeAccent5DataBar.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
