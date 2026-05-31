using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class DataLabelHyperlinkDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["A4"].PutValue("Item 3");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column that will hold the hyperlink text for each data point
                sheet.Cells["C1"].PutValue("Link");
                sheet.Hyperlinks.Add("C2", 1, 1, "https://www.example.com");
                sheet.Hyperlinks[0].TextToDisplay = "Visit";
                sheet.Hyperlinks.Add("C3", 1, 1, "https://www.example.org");
                sheet.Hyperlinks[1].TextToDisplay = "Visit";
                sheet.Hyperlinks.Add("C4", 1, 1, "https://www.example.net");
                sheet.Hyperlinks[2].TextToDisplay = "Visit";

                // Create a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels and link them to the hyperlink column (C2:C4)
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;               // Show the numeric value
                series.DataLabels.LinkedSource = "C2:C4";          // Link label text to hyperlink cells
                series.DataLabels.NumberFormatLinked = false;     // Keep number format independent

                // Save the workbook
                string outputPath = "DataLabelHyperlinkDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelHyperlinkDemo.Run();
        }
    }
}