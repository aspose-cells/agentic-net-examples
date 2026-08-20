// Title: Aspose.Cells C# – Add a Green Data Bar Conditional Format to Column L for Progress Percentages
// Description: This example creates a new workbook, populates column L with numeric progress values (0‑90%), and applies a DataBar conditional format with automatic minimum and maximum scaling. The bar is rendered in green and the cell value is displayed alongside the bar. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | DataBar conditional formatting | Excel progress bar | column L formatting | automatic min max | green data bar | save workbook as XLSX | conditional formatting API
// Common Searches: Aspose.Cells add data bar to a column | C# conditional formatting data bar example | how to show progress percentages with data bars in Excel using Aspose | set automatic min and max for data bar Aspose.Cells | change data bar color to green in Aspose.Cells
// Developer Intent: Generate an Excel file and visualize numeric progress values in column L using a green data‑bar conditional format.
// Use Cases: Create a project‑status report where each row displays a visual progress bar in column L. | Build a reusable Excel template that automatically scales data bars for any numeric range. | Export performance metrics to a dashboard‑style spreadsheet with instant visual cues.
// AI Prompts: Write C# code with Aspose.Cells to apply a green DataBar conditional format to column L, using automatic min/max values. | Show how to fill column L with percentage numbers and enable the DataBar to display the cell value. | Explain how to customize the DataBar color or gradient in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsDataBarExample
{
    // This example creates a new workbook, populates column L with numeric progress values (0‑90%), and applies a DataBar conditional format with automatic minimum and maximum scaling. The bar is rendered in green and the cell value is displayed alongside the bar. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column L (index 11) with sample progress percentages (0% to 100%)
            for (int row = 0; row < 10; row++)
            {
                // Example: 10% increments
                sheet.Cells[row, 11].PutValue(row * 10);
            }

            // Add an empty conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range for the data bar (column L, rows 0‑9)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 11,
                EndColumn = 11
            };
            cfCollection.AddArea(area);

            // Add a DataBar condition to the collection
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = cfCollection[conditionIndex];

            // Configure the DataBar properties
            DataBar dataBar = condition.DataBar;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum value based on data
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum value based on data
            dataBar.Color = Color.Green;                                   // Bar color
            dataBar.ShowValue = true;                                      // Show cell values alongside bars

            // Save the workbook to an XLSX file
            workbook.Save("ProgressDataBar.xlsx", SaveFormat.Xlsx);
        }
    }
}
