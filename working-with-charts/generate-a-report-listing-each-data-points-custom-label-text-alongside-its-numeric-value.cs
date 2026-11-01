// Title: Generate an Excel workbook that lists custom chart data label texts with their numeric values using Aspose.Cells for .NET
// AI Prompts: Create a column chart from worksheet data, assign custom text to each data point's label, and fill a new worksheet with two columns (label and original numeric value) before saving the file. | Rewrite the example so the numeric values are obtained from the chart series itself rather than the source cells when building the label‑value report. | Enhance the report sheet by inserting a total row that sums all numeric values and applies bold formatting to the summary.
// Common Searches: Aspose.Cells how to read custom data label text from chart points in C# | C# extract chart point labels and values to a separate worksheet using Aspose.Cells | Write a report sheet with custom chart labels and their numbers in Aspose.Cells .NET | Save Excel file with column chart having custom data labels and a summary row using Aspose.Cells
// Tags: extract chart point labels Aspose.Cells | write label-value report Excel Aspose.Cells | column chart with custom point text .NET | add totals row to worksheet Aspose.Cells | populate report sheet from chart series Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook with sample category/value data, builds a column chart with custom text labels for each point, generates a "Report" worksheet that lists each custom label alongside its numeric value, and saves the workbook as CustomLabelReport.xlsx.
class CustomLabelReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet for data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data (categories in column A, values in column B)
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("Item A");
            dataSheet.Cells["A3"].PutValue("Item B");
            dataSheet.Cells["A4"].PutValue("Item C");
            dataSheet.Cells["A5"].PutValue("Item D");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["B3"].PutValue(85);
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["B5"].PutValue(95);

            // Add a column chart based on the data
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = dataSheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Define custom label texts for each data point
            string[] customLabels = { "Alpha", "Beta", "Gamma", "Delta" };

            // Apply custom labels to each point
            for (int i = 0; i < customLabels.Length; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];

                // Show the value in the data label
                point.DataLabels.ShowValue = true;

                // Set custom text for the data label (using the Text property)
                point.DataLabels.Text = customLabels[i];
            }

            // Add a worksheet for the report
            Worksheet reportSheet = workbook.Worksheets.Add("Report");
            reportSheet.Cells["A1"].PutValue("Custom Label");
            reportSheet.Cells["B1"].PutValue("Numeric Value");

            // Extract custom label and numeric value from each point and write to the report sheet
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];

                // Retrieve custom label text (using the Text property)
                string label = point.DataLabels.Text;

                // Retrieve numeric value directly from the data sheet (column B, rows 2‑5)
                double value = dataSheet.Cells[i + 1, 1].DoubleValue; // zero‑based column index

                // Write to report sheet
                reportSheet.Cells[i + 1, 0].PutValue(label);   // Column A
                reportSheet.Cells[i + 1, 1].PutValue(value);  // Column B
            }

            // Save the workbook
            string outputPath = "CustomLabelReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
