// Title: C# – Create a Column Chart, Embed a Hidden Legend Tag, and Retrieve It with Aspose.Cells
// Description: This example shows how to generate a column chart in a new workbook, store a custom identifier in the chart's hidden title (used as a legend tag), save the file, reload it, and read the hidden title back as metadata using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart legend tag | store custom metadata in Excel chart | hidden chart title Aspose.Cells | retrieve chart tag .NET | column chart Excel automation | Excel chart custom identifier | Aspose.Cells chart metadata
// Common Searches: Aspose.Cells add hidden tag to chart legend | C# read hidden title from Excel chart | store custom string in Aspose.Cells chart | retrieve chart metadata after saving workbook | embed identifier in Excel chart using Aspose.Cells
// Developer Intent: Attach a custom identifier to a chart’s legend for later retrieval without affecting the visual layout.
// Use Cases: Assign unique IDs to charts for automated report processing. | Embed version or source information in a chart without visible changes. | Mark charts for selective extraction during batch analysis of Excel files.
// AI Prompts: Write C# code that creates a column chart and saves a hidden legend tag using Aspose.Cells. | Show how to load a workbook and extract the hidden title text from a chart to get the stored tag. | Suggest an alternative way to attach custom metadata to a chart object without using the title property.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to generate a column chart in a new workbook, store a custom identifier in the chart's hidden title (used as a legend tag), save the file, reload it, and read the hidden title back as metadata using Aspose.Cells for .NET.
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
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Store a custom tag using the chart's Title (hidden from view)
            chart.Title.Text = "MyCustomLegendTag";
            chart.Title.IsVisible = false;

            // Save the workbook
            string filePath = "ChartWithLegendTag.xlsx";
            workbook.Save(filePath);

            // Verify the file exists before loading
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The workbook file was not found.", filePath);

            // Load the workbook again to demonstrate retrieval of the tag
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Chart loadedChart = loadedSheet.Charts[0];

            // Retrieve the custom tag from the chart's Title text
            string retrievedTag = loadedChart.Title.Text;
            Console.WriteLine("Retrieved legend tag: " + (retrievedTag ?? "null"));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
