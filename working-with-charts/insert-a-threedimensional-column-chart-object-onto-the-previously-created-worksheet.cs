// Title: How to add a three‑dimensional column chart to an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a 3d column chart to an Excel file with Aspose.Cells C# | Aspose.Cells C# example for creating a three dimensional column chart | setting data source for a 3D column chart using Aspose.Cells .NET | Aspose.Cells chart title and series configuration in C# | saving an Excel workbook with a 3D column chart using Aspose.Cells
// Tags: Aspose.Cells create 3D column chart | C# set chart series range Aspose.Cells | Aspose.Cells define chart categories | Aspose.Cells position chart on worksheet | Aspose.Cells save workbook with chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills it with sample data, adds a three‑dimensional column chart at a specified location, assigns the series values and category labels, sets a chart title, and saves the file as ThreeDColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet where the chart will be placed
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare sample data for the chart
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            // -------------------------------------------------

            // Add a three‑dimensional column chart to the worksheet
            // Note: In some Aspose.Cells versions the Add method expects the ChartType as the first argument.
            int chartIdx = sheet.Charts.Add(ChartType.Column3D, 5, 0, 400, 600);
            Chart chart = sheet.Charts[chartIdx];

            // Define the data range for the series (values) and categories (labels)
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Optional: give the chart a title
            chart.Title.Text = "3‑D Column Chart";

            // Save the workbook with the chart embedded
            string outputPath = "ThreeDColumnChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
