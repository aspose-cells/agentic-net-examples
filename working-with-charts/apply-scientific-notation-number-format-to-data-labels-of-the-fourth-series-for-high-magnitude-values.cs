// Title: Aspose.Cells C# – Apply Scientific Notation to Data Labels of the Fourth Series in a Column Chart
// Description: This example creates a workbook with four high‑magnitude series, adds a column chart, enables data labels for the fourth series, and formats those labels using the scientific notation pattern "0.00E+00". The resulting file (ScientificNotationDataLabels.xlsx) demonstrates how to present large numbers clearly in Excel charts with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | column chart | data labels | scientific notation | number format | fourth series | high magnitude values | Excel export | chart formatting
// Common Searches: Aspose.Cells set scientific notation for chart data labels C# | format fourth series data labels as 0.00E+00 Aspose.Cells | apply number format to specific chart series .NET | display large numbers in Excel chart using Aspose.Cells
// Developer Intent: Format the data labels of the fourth series in a column chart to show values in scientific notation.
// Use Cases: Generate Excel reports where only the last series of a multi‑series column chart uses scientific notation for readability. | Create dashboards that highlight high‑value data points by applying a custom number format to selected chart series. | Automate workbook creation with Aspose.Cells where specific series require a different numeric display style.
// AI Prompts: Provide C# code with Aspose.Cells to set the data label format "0.00E+00" for the fourth series of a column chart. | Show how to enable data labels and apply scientific notation to a chosen series in an Aspose.Cells chart. | Explain how to programmatically change the number format of chart series data labels based on their index using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsScientificNotationDemo
{
    // This example creates a workbook with four high‑magnitude series, adds a column chart, enables data labels for the fourth series, and formats those labels using the scientific notation pattern "0.00E+00". The resulting file (ScientificNotationDataLabels.xlsx) demonstrates how to present large numbers clearly in Excel charts with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Columns B‑E: Four series with high magnitude values
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["E1"].PutValue("Series4");

            sheet.Cells["B2"].PutValue(1.2e6);
            sheet.Cells["B3"].PutValue(2.5e6);
            sheet.Cells["B4"].PutValue(3.8e6);
            sheet.Cells["B5"].PutValue(4.1e6);

            sheet.Cells["C2"].PutValue(5.0e6);
            sheet.Cells["C3"].PutValue(6.3e6);
            sheet.Cells["C4"].PutValue(7.7e6);
            sheet.Cells["C5"].PutValue(8.9e6);

            sheet.Cells["D2"].PutValue(9.2e6);
            sheet.Cells["D3"].PutValue(10.5e6);
            sheet.Cells["D4"].PutValue(11.8e6);
            sheet.Cells["D5"].PutValue(12.3e6);

            sheet.Cells["E2"].PutValue(13.6e6);
            sheet.Cells["E3"].PutValue(14.9e6);
            sheet.Cells["E4"].PutValue(15.2e6);
            sheet.Cells["E5"].PutValue(16.4e6);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the four series to the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].XValues = "A2:A5";

            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].XValues = "A2:A5";

            chart.NSeries.Add("D2:D5", true);
            chart.NSeries[2].XValues = "A2:A5";

            chart.NSeries.Add("E2:E5", true);
            chart.NSeries[3].XValues = "A2:A5";

            // Enable data labels for the fourth series (index 3)
            Series fourthSeries = chart.NSeries[3];
            fourthSeries.DataLabels.ShowValue = true;

            // Apply scientific notation format to the data labels
            // Format string "0.00E+00" displays numbers like 1.23E+06
            fourthSeries.DataLabels.NumberFormat = "0.00E+00";

            // Save the workbook
            workbook.Save("ScientificNotationDataLabels.xlsx");
        }
    }
}
