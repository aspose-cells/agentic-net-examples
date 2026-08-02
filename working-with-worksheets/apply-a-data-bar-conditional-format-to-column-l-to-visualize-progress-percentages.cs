// Title: Add a Green Data Bar to Column L for Progress Percentages with Aspose.Cells (.NET)
// Description: Creates a new workbook, fills column L (L1‑L10) with incremental progress values, and applies a DataBar conditional format with automatic 0‑100 % range, green bar color, and visible numeric values. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# data bar | conditional formatting column L | progress percentage data bar | Excel data bar Aspose.NET | automatic min max data bar | green data bar Aspose.Cells | show values with data bar
// Common Searches: how to add a data bar to a specific column using Aspose.Cells | Aspose.Cells example for visualizing progress percentages | set automatic min and max for data bar in .NET | display numeric values with data bar in Excel via Aspose
// Developer Intent: Generate an Excel file and apply a green data‑bar conditional format to column L to illustrate progress percentages.
// Use Cases: Project status reports where each task’s completion % appears as a green bar in column L. | Sales dashboards that show monthly target achievement with data bars for quick visual comparison. | Automated KPI sheets that highlight progress metrics using conditional data bars.
// AI Prompts: Write C# code using Aspose.Cells to add a green data‑bar conditional format to column L with automatic 0‑100 % limits and show the numeric values, then save as XLSX. | Provide an Aspose.Cells example that populates column L with percentage values and applies a data‑bar conditional format, displaying the values beside the bars. | Explain how to change the data‑bar color, range, or value visibility in an Aspose.Cells conditional formatting rule for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, fills column L (L1‑L10) with incremental progress values, and applies a DataBar conditional format with automatic 0‑100 % range, green bar color, and visible numeric values. The workbook is saved as an XLSX file.
    public class DataBarProgressInColumnL
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate column L (index 11) with sample progress percentages (0 to 100)
            for (int row = 0; row < 10; row++)
            {
                // Example: progress increases by 10% each row
                worksheet.Cells[row, 11].PutValue(row * 10);
            }

            // Add an empty conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

            // Define the range for the data bar (column L, rows 0‑9)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 11,   // Column L
                EndColumn = 11
            };
            cfCollection.AddArea(area);

            // Add a DataBar condition to the collection
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = cfCollection[conditionIndex];

            // Configure the DataBar properties
            DataBar dataBar = condition.DataBar;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum = 0%
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum = 100%
            dataBar.Color = Color.Green;                                     // Bar color
            dataBar.ShowValue = true;                                        // Show the numeric value alongside the bar

            // Save the workbook to an XLSX file
            workbook.Save("DataBarProgressColumnL.xlsx", SaveFormat.Xlsx);
        }
    }
}
