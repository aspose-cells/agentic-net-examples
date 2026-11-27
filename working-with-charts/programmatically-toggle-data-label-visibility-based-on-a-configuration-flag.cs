// Title: Toggle Excel chart data label visibility with Aspose.Cells for .NET (C#) using a command‑line flag
// Common Searches: Aspose.Cells C# hide chart data labels based on a boolean parameter | how to programmatically show or hide data labels in an Excel chart using Aspose.Cells | set DataLabels.IsDeleted property in Aspose.Cells chart from command line argument | C# example to conditionally display chart values in Aspose.Cells workbook | toggle column chart data labels in Aspose.Cells .NET with runtime flag
// Tags: Aspose.Cells chart data label toggle | DataLabels.IsDeleted property C# | conditional chart labeling Aspose.Cells | Excel column chart label visibility .NET | command line flag chart customization Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DataLabelToggleDemo
{
    // The sample reads a boolean flag from the command line, creates a workbook with sample data, adds a column chart, and uses the Series.DataLabels.IsDeleted property to hide or show data labels. When enabled, it also displays the values, sets the label position, and customizes the font before saving the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the flag from command‑line arguments.
            bool showDataLabels = false; // default: hide
            if (args.Length > 0)
            {
                if (!bool.TryParse(args[0], out showDataLabels))
                {
                    Console.WriteLine("Invalid argument. Expected 'true' or 'false'. Using default (false).");
                    showDataLabels = false;
                }
            }

            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series.
                Series series = chart.NSeries[0];

                // Configure data labels based on the flag.
                DataLabels dataLabels = series.DataLabels;

                // Show or hide the whole data‑label collection.
                // DataLabels.IsDeleted = true hides the labels; false shows them.
                dataLabels.IsDeleted = !showDataLabels;

                // When showing, also enable the value display.
                if (showDataLabels)
                {
                    dataLabels.ShowValue = true;
                    // Optional: customize appearance.
                    dataLabels.Position = LabelPositionType.InsideEnd;
                    dataLabels.Font.Size = 10;
                    dataLabels.Font.Color = System.Drawing.Color.DarkBlue;
                }

                // Save the workbook.
                string outputPath = "DataLabelToggleResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'. Data labels are {(showDataLabels ? "visible" : "hidden")}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
