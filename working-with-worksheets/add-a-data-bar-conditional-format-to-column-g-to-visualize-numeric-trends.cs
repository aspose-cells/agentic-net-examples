// Title: Apply a Blue Data Bar Conditional Format to Column G with Aspose.Cells for C#
// Description: This example creates a new workbook, fills cells G1‑G10 with numeric values, and adds a conditional‑formatting rule that displays a blue data bar. The data bar uses automatic minimum and maximum values, shows the cell values, and the workbook is saved as DataBarColumnG.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel | Data Bar | Conditional Formatting | Column G | Automatic Min Max | Blue data bar | Workbook example
// Common Searches: Aspose.Cells data bar column G C# | How to add conditional formatting data bar in Aspose.Cells .NET | C# code for Excel data bar conditional format | Set data bar color Aspose.Cells | Automatic minimum maximum data bar Aspose.Cells
// Developer Intent: Add a data‑bar conditional format to column G to visualize numeric trends in an Excel file generated with Aspose.Cells.
// Use Cases: Show sales figures in column G with blue bars for quick performance comparison. | Visualize project completion percentages in column G using data bars. | Create an inventory‑level heat map in column G with proportional bars. | Display KPI scores in column G for an executive dashboard. | Highlight financial variance in column G with a visual bar indicator.
// AI Prompts: Generate C# Aspose.Cells code to apply a red data bar to column H with custom minimum 0 and maximum 100. | Show how to modify the example to hide cell values while keeping the blue data bar visible. | Provide a sample that adds both a data bar and a three‑color scale to column G in the same worksheet. | Write a reusable method that applies a data‑bar conditional format to any specified column range. | Create an Aspose.Cells script that uses a gradient color data bar and sets the bar direction to left‑to‑right.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsDataBarExample
{
    // This example creates a new workbook, fills cells G1‑G10 with numeric values, and adds a conditional‑formatting rule that displays a blue data bar. The data bar uses automatic minimum and maximum values, shows the cell values, and the workbook is saved as DataBarColumnG.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column G (index 6) with sample numeric data
            for (int row = 0; row < 10; row++)
            {
                sheet.Cells[row, 6].PutValue(row * 10 + 5); // G1..G10
            }

            // Add an empty conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range for the data bar (column G, rows 0-9)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 6,
                EndColumn = 6
            };
            cfCollection.AddArea(area);

            // Add a DataBar condition to the collection
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = cfCollection[conditionIndex];

            // Configure the DataBar properties
            DataBar dataBar = condition.DataBar;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum based on data
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum based on data
            dataBar.Color = Color.Blue;                                   // Bar color
            dataBar.ShowValue = true;                                     // Show cell values

            // Save the workbook to a file
            workbook.Save("DataBarColumnG.xlsx", SaveFormat.Xlsx);
        }
    }
}
