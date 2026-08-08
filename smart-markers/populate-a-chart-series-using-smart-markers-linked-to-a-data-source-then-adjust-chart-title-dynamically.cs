// Title: C# – Populate an Excel column chart from smart markers and set a dynamic title with Aspose.Cells
// Description: This example demonstrates how to create a workbook, define smart markers for product and unit columns, expand them with a List<dynamic> data source using WorkbookDesigner, add a column chart, bind the NSeries and category data to the populated cells, set the chart title from the first product cell, recalculate the chart, and save the file as an .xlsx document.
// Keywords: Aspose.Cells | C# Excel chart | smart markers | dynamic chart title | NSeries range | WorkbookDesigner data source | populate chart from cells | Excel column chart automation
// Common Searches: Aspose.Cells smart markers fill chart data C# | Set Excel chart title from cell value using Aspose.Cells | Link NSeries to expanded smart marker range | Create column chart after processing smart markers | Dynamic chart title Aspose.Cells .NET
// Developer Intent: Generate a column chart whose series and categories are automatically filled by smart markers and update the chart title based on the first populated product name.
// Use Cases: Automated sales dashboards that insert product‑unit data via smart markers and visualize it in a column chart. | Templates that reuse the same workbook layout while swapping data sources, producing charts with titles that reflect the current dataset. | Reporting tools that need to adjust chart titles on‑the‑fly without recreating the chart object.
// AI Prompts: Write C# code with Aspose.Cells to create a column chart from smart‑marker‑expanded cells and set the title using the first product name. | Explain how to bind NSeries and CategoryData to ranges that are populated by WorkbookDesigner smart markers and ensure the chart updates correctly. | Show how to change an Excel chart title dynamically after processing smart markers without rebuilding the chart.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// This example demonstrates how to create a workbook, define smart markers for product and unit columns, expand them with a List<dynamic> data source using WorkbookDesigner, add a column chart, bind the NSeries and category data to the populated cells, set the chart title from the first product cell, recalculate the chart, and save the file as an .xlsx document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Units");

            // Smart marker row – will be expanded by WorkbookDesigner
            sheet.Cells["A2"].PutValue("&=$Product");
            sheet.Cells["B2"].PutValue("&=$Units");

            // Define the smart‑marker range (required by the designer)
            AsposeRange smartRange = sheet.Cells.CreateRange("A2:B2");
            smartRange.Name = "_CellsSmartMarkers";

            // Sample data source for the smart markers
            var products = new List<dynamic>
            {
                new { Product = "Apple",  Units = 120 },
                new { Product = "Banana", Units = 80  },
                new { Product = "Cherry", Units = 150 }
            };

            // Set the data source and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", products);
            designer.Process(); // populates A2:B4 with the data above

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Populate the series using NSeries (linked to the filled cells)
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Dynamically set the chart title based on the first product name
            string firstProduct = sheet.Cells["A2"].StringValue;
            chart.Title.Text = $"Units Sold – {firstProduct}";
            chart.Title.OverLay = true; // overlay title without resizing the chart

            // Ensure the chart reflects the latest data
            chart.Calculate();

            // Save the workbook
            string outputPath = "SmartMarkersChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
