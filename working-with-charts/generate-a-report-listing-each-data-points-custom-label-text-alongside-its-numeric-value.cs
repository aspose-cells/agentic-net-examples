// Title: Create an Excel workbook with a column chart that uses custom data‑label text and a separate worksheet listing each label with its numeric value using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that builds a column chart, disables automatic label text, assigns a custom label (e.g., "Item: <category>") to each point, and populates a new worksheet with the label and its Y‑value. | Update an Aspose.Cells chart example to generate a summary sheet that extracts every point's custom DataLabels.Text and corresponding numeric value, then save the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# set custom text for each chart point label | How to export chart point labels and values to a worksheet with Aspose.Cells | Create a report sheet of custom data‑labels and numeric values from an Excel chart in .NET | C# Aspose.Cells column chart with custom data labels and summary worksheet
// Tags: custom data labels Aspose.Cells chart | export chart point values to worksheet C# | generate label/value report worksheet .NET | set point.DataLabels.Text Aspose.Cells | column chart custom labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomLabelReport
{
    // The program creates a new workbook, adds a column chart with custom data‑label text derived from the category column, disables automatic label generation, and then writes each custom label together with its numeric Y‑value to a separate "Report" worksheet before saving the file as CustomLabelReport.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare data for the chart (first worksheet)
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");

                // Sample data
                dataSheet.Cells["A2"].PutValue("Alpha");
                dataSheet.Cells["A3"].PutValue("Beta");
                dataSheet.Cells["A4"].PutValue("Gamma");

                dataSheet.Cells["B2"].PutValue(1500);
                dataSheet.Cells["B3"].PutValue(2750);
                dataSheet.Cells["B4"].PutValue(3200);

                // -------------------------------------------------
                // 2. Add a chart and enable custom data labels
                // -------------------------------------------------
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = dataSheet.Charts[chartIndex];

                // Bind data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // show numeric value
                series.DataLabels.ShowCategoryName = false; // we will use custom text

                // Assign custom label to each point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Disable auto‑generated text so we can set our own
                    point.DataLabels.IsAutoText = false;

                    // Retrieve the category name from the source data (column A)
                    string category = dataSheet.Cells[i + 1, 0].StringValue; // A2, A3, ...

                    // Set custom label
                    point.DataLabels.Text = $"Item: {category}";
                }

                // -------------------------------------------------
                // 3. Generate a report worksheet listing label text and value
                // -------------------------------------------------
                Worksheet reportSheet = workbook.Worksheets.Add("Report");
                // Headers
                reportSheet.Cells["A1"].PutValue("Custom Label");
                reportSheet.Cells["B1"].PutValue("Numeric Value");

                // Populate rows with data from the chart points
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    int row = i + 2; // start from row 2 (1‑based index)

                    // Custom label text
                    reportSheet.Cells[row, 0].PutValue(point.DataLabels.Text);

                    // Numeric value (YValue)
                    reportSheet.Cells[row, 1].PutValue(point.YValue);
                }

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("CustomLabelReport.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
