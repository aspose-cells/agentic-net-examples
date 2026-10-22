// Title: Add a column chart, assign a custom tag to its legend, and read the tag back from a saved .xlsx using Aspose.Cells for .NET
// AI Prompts: Write C# using Aspose.Cells to create a column chart, assign a unique key to the legend's Text property, save the workbook, reload it, and display the key. | Show how to embed a developer‑defined value in a chart legend via Legend.Text and later extract it from the .xlsx file with Aspose.Cells. | Provide a concise example that stores arbitrary metadata in a chart legend, persists the workbook, and reads the metadata back in a .NET application.
// Common Searches: aspocells add custom identifier to chart legend and retrieve after save | c# Aspose.Cells store value in legend text of column chart | how to read legend Text property from a saved Excel file using Aspose.Cells | embedding metadata in Excel chart legend with Aspose.Cells for .NET | retrieve custom legend tag from .xlsx using Aspose.Cells C#
// Tags: Aspose.Cells chart legend custom identifier | C# column chart legend Text property | store metadata in Excel chart legend | retrieve legend value from saved workbook | Aspose.Cells .xlsx legend persistence

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds sample data, inserts a column chart, sets the legend position to bottom, stores a custom tag in the legend's Text property, saves the file as .xlsx, reloads the workbook, and reads back the legend tag.
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
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the legend and assign a custom tag (using Text as a placeholder)
            Legend legend = chart.Legend;
            legend.Position = LegendPositionType.Bottom;
            legend.Text = "MyCustomLegendTag";

            // Save the workbook
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithLegendTag.xlsx");
            workbook.Save(filePath);

            // ------------------------------
            // Load the workbook and retrieve the tag
            // ------------------------------
            if (File.Exists(filePath))
            {
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Chart loadedChart = loadedSheet.Charts[0];
                Legend loadedLegend = loadedChart.Legend;

                // Retrieve the custom tag (stored in Text)
                string retrievedTag = loadedLegend.Text;

                // Output the retrieved tag
                Console.WriteLine("Retrieved Legend Tag: " + retrievedTag);
            }
            else
            {
                Console.WriteLine("Error: The file was not found after saving.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
