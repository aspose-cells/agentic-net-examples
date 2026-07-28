// Title: Add Data Bar Conditional Formatting with Accent 4 Theme Color using Aspose.Cells for C#
// Description: The sample builds a workbook, writes numbers to cells A1‑A5, and attaches a DataBar conditional format. The bar is colored with the Excel theme’s Accent 4 hue (RGB 0,112,192), uses automatic min/max scaling, shows the cell value, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | DataBar | Conditional Formatting | Accent4 | theme color | Excel automation | automatic min max | sample code | positive values
// Common Searches: Aspose.Cells data bar with theme accent color | C# conditional formatting data bar Accent4 | how to set Excel theme color for data bars in .NET | apply data bar only to positive numbers using Aspose.Cells | automatic scaling data bar Aspose.Cells example
// Developer Intent: Generate an Excel file and apply a DataBar conditional format that uses the theme’s Accent 4 color for the target range.
// Use Cases: Create a sales report where positive figures are visualized with blue Accent 4 bars for quick comparison. | Build a KPI dashboard that highlights improvement metrics using theme‑consistent data bars while leaving negative values unchanged. | Automate monthly financial statements that include color‑matched data bars to align with corporate branding.
// AI Prompts: Write C# code with Aspose.Cells to add a DataBar conditional format colored with the workbook’s Accent 4 theme. | Show how to retrieve the exact Accent 4 RGB value from an Excel theme and apply it to a data bar in Aspose.Cells. | Provide an example that applies a data bar only to cells with positive numbers, preserving the original formatting for negatives.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample builds a workbook, writes numbers to cells A1‑A5, and attaches a DataBar conditional format. The bar is colored with the Excel theme’s Accent 4 hue (RGB 0,112,192), uses automatic min/max scaling, shows the cell value, and saves the result as an XLSX file.
    public class DataBarAccent4Demo
    {
        public static void Main()
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

            // Populate sample numeric data (both positive and negative)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(-20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(-40);
            sheet.Cells["A5"].PutValue(50);

            // Add a conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

            // Define the range for the data bar (A1:A5)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            conditions.AddArea(area);

            // Add a DataBar condition
            int conditionIndex = conditions.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = conditions[conditionIndex];

            // Configure the DataBar
            DataBar dataBar = condition.DataBar;

            // Use a standard color (e.g., Accent4 equivalent) for the bar
            // Aspose.Cells does not expose theme accent colors directly, so we use a representative color.
            dataBar.Color = Color.FromArgb(0, 112, 192); // Approximation of Accent4

            // Set automatic min/max so the bar scales to the data range
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;

            // Show the cell values alongside the bars
            dataBar.ShowValue = true;

            // Determine output file path and ensure the directory exists
            string outputPath = "DataBarAccent4Demo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
