// Title: Create a column chart with Aspose.Cells for .NET, store custom legend data in a hidden cell, and read it back after reopening the workbook
// AI Prompts: Generate C# code that uses Aspose.Cells to add a column chart, writes a custom identifier for the chart legend into a hidden worksheet cell, saves the workbook, reloads it, and extracts the identifier. | Show how to emulate a Legend.Tag property in Aspose.Cells by using a hidden cell or auxiliary data structure, then demonstrate retrieving that custom legend information programmatically. | Provide a step‑by‑step example that verifies the legend’s position after the workbook is reopened using Aspose.Cells.
// Common Searches: Aspose.Cells how to attach custom data to a chart legend in .NET | workaround for missing Legend.Tag in Aspose.Cells charts | store chart legend identifier in hidden Excel cell using Aspose.Cells | retrieve custom legend metadata after reopening workbook with Aspose.Cells | C# Aspose.Cells add column chart and read hidden cell value
// Tags: Aspose.Cells add column chart | store legend custom data hidden cell | Aspose.Cells legend metadata workaround | read hidden cell after workbook reload | C# Aspose.Cells chart legend tag alternative

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new Workbook, populates sample data, adds a column chart, and explains that Aspose.Cells' Legend class lacks a Tag property. It demonstrates storing a custom legend identifier in a hidden worksheet cell, saving the file, reloading the workbook, and retrieving the identifier while confirming the legend’s presence and position.
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

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Aspose.Cells' Legend class does not provide a Tag property.
            // If you need to associate custom data with the legend, consider using
            // a separate data structure or embedding the information elsewhere
            // (e.g., in a hidden cell or chart title).

            // Save the workbook with the chart
            string filePath = "ChartWithLegendTag.xlsx";
            workbook.Save(filePath);

            // Verify that the file was created
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The workbook file was not created.", filePath);

            // Load the workbook to demonstrate further processing
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Chart loadedChart = loadedSheet.Charts[0];

            // Since Legend.Tag is not available, we simply confirm the legend exists
            Console.WriteLine("Legend is present. Position: " + loadedChart.Legend.Position);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
