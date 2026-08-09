// Title: Add a column chart series from an in‑memory double[] using Aspose.Cells for .NET (C#)
// Description: Shows how to import a double[] into a worksheet, create a column chart, build the data range address, add the series with NSeries.Add (vertical orientation), and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# | chart series from array | ImportArray double | NSeries.Add | column chart | in‑memory data | Excel export | .NET
// Common Searches: Aspose.Cells add chart series from double array | Import double[] into Excel chart C# | Create column chart from in‑memory data Aspose.Cells | NSeries.Add vertical data Aspose.Cells | Plot runtime array without writing cells
// Developer Intent: Create a chart series directly from a double[] at runtime, avoiding manual cell‑by‑cell entry.
// Use Cases: Generate a sales‑trend column chart from a calculation that returns a double[] | Visualize simulation results stored in a double[] immediately in an Excel chart | Produce a quick sensor‑reading report by importing a double[] and displaying it as a column chart
// AI Prompts: How can I add multiple series from several double[] arrays to the same chart with Aspose.Cells? | Show code to assign custom category labels from a string[] to a series created from a double[] array. | Explain how to format the column chart (colors, axis titles, data labels) after adding a series from an in‑memory array.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesFromArray
{
    // Shows how to import a double[] into a worksheet, create a column chart, build the data range address, add the series with NSeries.Add (vertical orientation), and save the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // In‑memory double array that will become the series values
            double[] values = new double[] { 12.5, 23.8, 35.0, 47.6, 58.9 };

            // Import the array vertically starting at cell A1 (row 0, column 0)
            // The ImportArray(double[], int, int, bool) rule is used here
            sheet.Cells.ImportArray(values, 0, 0, true);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Build the address of the imported data range (A1:A5)
            string dataRange = $"=Sheet1!A1:A{values.Length}";

            // Add the series to the chart using the NSeries.Add(string, bool) rule
            // The second parameter 'true' indicates that the data is arranged vertically
            chart.NSeries.Add(dataRange, true);

            // (Optional) Set category labels if desired – here we use simple numeric categories
            // chart.NSeries.CategoryData = $"=Sheet1!B1:B{values.Length}";

            // Save the workbook
            workbook.Save("ChartFromDoubleArray.xlsx");
        }
    }
}
