// Title: Label chart points below the average as "Low" using LINQ with Aspose.Cells for .NET
// Common Searches: how to label chart points below average as Low using Aspose.Cells C# | Aspose.Cells LINQ filter chart series points under average | set custom data label text for low values in a column chart Aspose.Cells | calculate average of chart series and modify point labels with LINQ in .NET | programmatically hide numeric values and show custom text for specific chart points Aspose.Cells
// Tags: filter chart points with LINQ Aspose.Cells | custom data label text for low values .NET | calculate series average Aspose.Cells | conditional chart point labeling Aspose.Cells | modify chart point data labels programmatically C# | column chart data label formatting Aspose.Cells

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsLinqLabelDemo
{
    // The example creates a workbook, populates categories and numeric values, adds a column chart, and forces chart calculation. Using LINQ it computes the average Y‑value of the series, selects points whose values are below that average, and updates each selected point's data label to display the custom text "Low" in bold red while hiding the original numeric value. The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: categories in column A and numeric values in column B
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["A6"].PutValue("E");

                sheet.Cells["B2"].PutValue(12);
                sheet.Cells["B3"].PutValue(25);
                sheet.Cells["B4"].PutValue(8);
                sheet.Cells["B5"].PutValue(30);
                sheet.Cells["B6"].PutValue(15);

                // Add a column chart that uses the data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B6", true);          // Y values
                chart.NSeries.CategoryData = "A2:A6";      // X categories

                // Ensure the chart is calculated so that point objects are populated
                chart.Calculate();

                // Get the first series (there is only one)
                Series series = chart.NSeries[0];

                // Cast the non‑generic collection to IEnumerable<ChartPoint>
                var points = series.Points.Cast<ChartPoint>();

                // Compute the average of the Y values of all points using LINQ
                double average = points.Average(p => Convert.ToDouble(p.YValue));

                Console.WriteLine($"Average value: {average}");

                // Select points whose Y value is below the average
                var lowPoints = points.Where(p => Convert.ToDouble(p.YValue) < average);

                // Set the data label of each low point to the text "Low"
                foreach (ChartPoint point in lowPoints)
                {
                    // Ensure the data label is visible
                    point.DataLabels.ShowValue = false;   // hide numeric value
                    point.DataLabels.Text = "Low";        // custom text
                    point.DataLabels.Font.Color = Color.Red;
                    point.DataLabels.Font.IsBold = true;
                }

                // Save the workbook to an XLSX file
                string outputPath = "LinqLowPointsDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
