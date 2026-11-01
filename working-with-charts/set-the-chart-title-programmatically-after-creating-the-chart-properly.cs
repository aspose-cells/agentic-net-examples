// Title: Set a column chart title programmatically with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart from a worksheet range and assigns a custom, visible title using Aspose.Cells. | Show how to configure the Title object of an Aspose.Cells chart and ensure it appears after saving the workbook. | Provide a step‑by‑step example that adds data, creates a column chart, sets the series range, and sets the chart title in Aspose.Cells.
// Common Searches: aspnet set chart title Aspose.Cells column chart example | C# Aspose.Cells how to display chart title in Excel file | programmatically add title to column chart using Aspose.Cells for .NET | Aspose.Cells chart title not visible after saving workbook | sample code for setting chart title in Aspose.Cells C#
// Tags: Aspose.Cells column chart title C# | set chart title Aspose.Cells .NET | Aspose.Cells chart series range example | create Excel chart with title using Aspose.Cells | Aspose.Cells workbook save with chart title | C# Aspose.Cells chart title visibility

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:B4 with month and sales data, adds a column chart, assigns the values range B2:B4 to the series, sets the chart title to "Quarterly Sales" and makes it visible, then saves the workbook as ChartWithTitle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series (values)
            chart.NSeries.Add("B2:B4", true);
            // Category (X‑axis) data is taken from the first column by default,
            // so explicit CategoryData assignment is not required for this version.

            // Set the chart title
            chart.Title.Text = "Quarterly Sales";
            chart.Title.IsVisible = true;

            // Define output file path
            string outputPath = "ChartWithTitle.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
