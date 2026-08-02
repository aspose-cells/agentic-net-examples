// Title: C# – List Custom Chart Labels and Values in a Separate Worksheet using Aspose.Cells
// Description: This C# example creates a workbook, adds sample data, builds a column chart, assigns custom text labels to each data point by disabling auto‑text, and then generates a "Report" worksheet that records each point’s label together with its numeric Y‑value, finally saving the file as CustomLabelReport.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel chart data labels | custom chart labels | disable auto text Aspose.Cells | extract chart point values | column chart report | generate worksheet from chart | YValue retrieval | chart point custom text
// Common Searches: Aspose.Cells get custom chart label value | C# export chart data labels to another sheet | disable auto‑generated data labels Aspose.Cells | retrieve YValue of chart points in Aspose.Cells | create report worksheet from chart data C#
// Developer Intent: Create an Excel workbook that extracts each chart point’s custom label and its numeric value into a dedicated report sheet.
// Use Cases: Audit chart data by providing a printable list of user‑defined labels and their values. | Feed chart point details into downstream reporting or BI tools that require tabular data. | Generate a summary sheet for presentations where chart labels differ from source categories.
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart, sets custom text for each data point, and creates a new worksheet listing those labels with their numeric values. | Explain how to turn off auto‑generated data label text for chart points and retrieve each point’s YValue using Aspose.Cells. | Provide step‑by‑step instructions to generate a report sheet from chart data labels and values, then save the workbook as an .xlsx file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomLabelReport
{
    // This C# example creates a workbook, adds sample data, builds a column chart, assigns custom text labels to each data point by disabling auto‑text, and then generates a "Report" worksheet that records each point’s label together with its numeric Y‑value, finally saving the file as CustomLabelReport.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data (Category and Value)
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["A2"].PutValue("Item 1");
                dataSheet.Cells["A3"].PutValue("Item 2");
                dataSheet.Cells["A4"].PutValue("Item 3");
                dataSheet.Cells["B2"].PutValue(150);
                dataSheet.Cells["B3"].PutValue(300);
                dataSheet.Cells["B4"].PutValue(450);

                // Add a column chart to the same sheet
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
                Chart chart = dataSheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;        // Show numeric values
                series.DataLabels.ShowCategoryName = false;

                // Assign a custom label to each data point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    // Disable auto‑generated text so we can set our own
                    point.DataLabels.IsAutoText = false;
                    // Example custom label: "Label_i"
                    point.DataLabels.Text = $"Label_{i + 1}";
                }

                // Create a new worksheet to hold the report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "Report";

                // Write headers
                reportSheet.Cells["A1"].PutValue("Custom Label");
                reportSheet.Cells["B1"].PutValue("Numeric Value");

                // Populate the report with each point's custom label and its value
                int reportRow = 1; // zero‑based index; row 1 is the second row (after header)
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    string label = point.DataLabels.Text;                     // custom label we set earlier
                    double value = Convert.ToDouble(point.YValue);           // numeric value of the point

                    reportSheet.Cells[reportRow, 0].PutValue(label); // Column A
                    reportSheet.Cells[reportRow, 1].PutValue(value); // Column B
                    reportRow++;
                }

                // Define output file path
                string outputPath = "CustomLabelReport.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
