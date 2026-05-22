using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class LinkDataLabelNumberFormatDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate sample data
            // -------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["E1"].PutValue("Formatted 1");
            sheet.Cells["F1"].PutValue("Formatted 2");
            sheet.Cells["G1"].PutValue("Formatted 3");

            // Category labels
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Raw numeric values for each series
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(2000);

            sheet.Cells["C2"].PutValue(1200);
            sheet.Cells["C3"].PutValue(1800);
            sheet.Cells["C4"].PutValue(2400);

            sheet.Cells["D2"].PutValue(1400);
            sheet.Cells["D3"].PutValue(2100);
            sheet.Cells["D4"].PutValue(2800);

            // Formatted strings that we want the data labels to follow
            sheet.Cells["E2"].PutValue("1,000 units");
            sheet.Cells["E3"].PutValue("1,500 units");
            sheet.Cells["E4"].PutValue("2,000 units");

            sheet.Cells["F2"].PutValue("1,200 units");
            sheet.Cells["F3"].PutValue("1,800 units");
            sheet.Cells["F4"].PutValue("2,400 units");

            sheet.Cells["G2"].PutValue("1,400 units");
            sheet.Cells["G3"].PutValue("2,100 units");
            sheet.Cells["G4"].PutValue("2,800 units");

            // -------------------------------------------------
            // Add a column chart
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add three series, each using its own numeric column
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.Add("D2:D4", true); // Series 3

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Link each series' data label number format to its formatted source column
            // -------------------------------------------------
            // Series 0 -> formatted values in column E
            Series s0 = chart.NSeries[0];
            s0.DataLabels.ShowValue = true;
            s0.DataLabels.LinkedSource = "E2:E4";
            s0.DataLabels.NumberFormatLinked = true;

            // Series 1 -> formatted values in column F
            Series s1 = chart.NSeries[1];
            s1.DataLabels.ShowValue = true;
            s1.DataLabels.LinkedSource = "F2:F4";
            s1.DataLabels.NumberFormatLinked = true;

            // Series 2 -> formatted values in column G
            Series s2 = chart.NSeries[2];
            s2.DataLabels.ShowValue = true;
            s2.DataLabels.LinkedSource = "G2:G4";
            s2.DataLabels.NumberFormatLinked = true;

            // Optional: style the data labels for better visibility
            foreach (Series ser in chart.NSeries)
            {
                ser.DataLabels.Font.Color = Color.DarkBlue;
                ser.DataLabels.Font.Size = 10;
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "LinkDataLabelNumberFormatDemo.xlsx";

            // Ensure we can write the file (overwrite if exists)
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            workbook.Save(outputPath);
        }
    }
}