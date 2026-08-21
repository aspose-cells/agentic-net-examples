// Title: Aspose.Cells C# – Apply Data Bar Conditional Formatting with Accent 4 Theme Color
// Description: Demonstrates how to create a workbook, fill cells A1:A5 with numeric data, add a DataBar conditional format, set the bar color to the theme's Accent 4 (RGB 0,112,192), use automatic minimum and maximum values, display the cell value, and save the file as an XLSX document.
// Keywords: Aspose.Cells | C# | DataBar | Conditional Formatting | Accent4 | Theme color | positive values | automatic min max | Excel XLSX | cell data bar color
// Common Searches: Aspose.Cells data bar Accent4 color C# example | how to set theme accent color for data bar in Aspose.Cells | conditional formatting data bar positive values Aspose.Cells .NET | apply automatic min and max to data bar using Aspose.Cells | C# code for data bar conditional format in Excel workbook
// Developer Intent: Add a DataBar conditional format that uses the workbook’s Accent 4 theme color for positive numbers in a specified range.
// Use Cases: Show sales performance with blue Accent 4 data bars in a financial report. | Highlight KPI values on a dashboard worksheet using theme‑based data bars. | Create a project‑status spreadsheet where each task’s progress is visualized with Accent 4 bars.
// AI Prompts: Generate C# Aspose.Cells code to apply a DataBar conditional format with the Accent 4 theme color to range B2:B12, displaying only positive values. | Provide an example that sets DataBar.Color to Theme.Accent4, uses automatic min/max, and shows the cell value in Aspose.Cells. | Explain how to retrieve the exact RGB value of Accent 4 from a workbook’s theme and apply it to a DataBar condition in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, fill cells A1:A5 with numeric data, add a DataBar conditional format, set the bar color to the theme's Accent 4 (RGB 0,112,192), use automatic minimum and maximum values, display the cell value, and save the file as an XLSX document.
    public class DataBarAccent4Demo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (both positive and negative values)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(-5);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(-15);
            sheet.Cells["A5"].PutValue(30);

            // Add a conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range for the data bar (A1:A5)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Add a DataBar condition
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = cfCollection[conditionIndex];

            // Configure the DataBar
            DataBar dataBar = condition.DataBar;
            // Use an approximate Accent4 color (since Theme.Accent4 is not directly accessible)
            dataBar.Color = Color.FromArgb(0, 112, 192); // Typical Accent4 blue
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            dataBar.ShowValue = true;

            // Save the workbook
            string outputPath = "DataBarAccent4Demo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
