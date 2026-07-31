// Title: C# Aspose.Cells Example – Set "#,##0" Number Format on Chart Data Labels (Thousand Separator, No Decimals)
// Description: This .NET sample builds a workbook, inserts category and quantity rows, creates a column chart, turns on data‑label values, and applies the number format "#,##0" so the labels show commas as thousand separators and omit decimal places, then saves the file as NumberFormatWithThousandSeparator.xlsx.
// Keywords: Aspose.Cells | C# .NET | custom number format | thousand separator | #,##0 | chart data labels | column chart | Excel export | format numbers in chart | Aspose.Cells number formatting
// Common Searches: Aspose.Cells set number format for chart labels | C# apply thousand separator to Excel chart data labels | How to use "#,##0" format in Aspose.Cells | format chart values without decimals Aspose.Cells .NET | custom number format for series data labels C#
// Developer Intent: The developer wants chart data‑label values to display with commas as thousand separators and without any decimal digits.
// Use Cases: Financial summary workbook where column chart labels need comma‑separated amounts for clarity. | Inventory report exporting to Excel with readable quantity labels on charts. | Sales dashboard that presents large numbers in charts without fractional parts. | Automated reporting tool that standardizes number appearance across generated spreadsheets.
// AI Prompts: Generate C# code using Aspose.Cells to apply the "#,##0" format to column chart data labels. | Show how to enable data labels and set a thousand‑separator number format in an Aspose.Cells workbook. | Explain the steps to customize chart label formatting in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNumberFormatDemo
{
    // This .NET sample builds a workbook, inserts category and quantity rows, creates a column chart, turns on data‑label values, and applies the number format "#,##0" so the labels show commas as thousand separators and omit decimal places, then saves the file as NumberFormatWithThousandSeparator.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["B2"].PutValue(1234);
            worksheet.Cells["B3"].PutValue(56789);
            worksheet.Cells["B4"].PutValue(101112);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and apply custom number format with thousand separators, no decimals
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.NumberFormat = "#,##0";

            // Save the workbook
            workbook.Save("NumberFormatWithThousandSeparator.xlsx");
        }
    }
}
