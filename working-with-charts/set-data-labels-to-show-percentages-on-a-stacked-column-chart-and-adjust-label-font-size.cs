// Title: Aspose.Cells for .NET – Add Percentage Data Labels and Set Font Size on a Stacked Column Chart
// Description: Creates a workbook, inserts sample data, adds a stacked column chart, enables percentage data labels for each series, and sets the label font size to 12 pt before saving the file as an Excel workbook.
// Keywords: Aspose.Cells stacked column chart | percentage data labels C# | chart label font size .NET | Aspose.Cells chart customization | Excel stacked column percentages | C# Aspose.Cells example
// Common Searches: Aspose.Cells show percentages on stacked column chart | change chart data label font size with Aspose.Cells | enable data labels for each series in Aspose.Cells | C# code for stacked column chart percentages | Aspose.Cells chart label styling
// Developer Intent: Add a stacked column chart to an Excel workbook and configure its data labels to display percentages with a custom font size using Aspose.Cells for .NET.
// Use Cases: Sales report that visualizes product contribution per quarter with percentage labels for quick insight. | Financial dashboard showing expense categories as stacked columns, each segment labeled with its share of the total. | Project management tracker where task completion ratios are displayed as percentages on stacked column charts for clear status communication.
// AI Prompts: Generate C# code with Aspose.Cells that creates a stacked column chart, shows percentage data labels, and sets the label font size to 12 points. | Provide an Aspose.Cells example that populates data, adds a stacked column chart, and customizes data label appearance (percentage display and font size). | Explain how to use ShowPercentage and Font.Size properties to format chart data labels in a .NET application using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, inserts sample data, adds a stacked column chart, enables percentage data labels for each series, and sets the label font size to 12 pt before saving the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked column chart to the worksheet
            // Correct enum value for stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data range (both products) and category data
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for each series, show percentages, and set font size
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowPercentage = true;   // Display percentage values
                series.DataLabels.Font.Size = 12;          // Adjust label font size
            }

            // Define output file path
            string outputPath = "StackedColumnDataLabels.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
