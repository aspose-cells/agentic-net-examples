// Title: Batch add a .crtx chart template to multiple Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a .crtx file into a byte array, iterates over a collection of .xlsx files, inserts the chart template into the first worksheet of each workbook, assigns a custom chart title, and saves the workbooks. | Show how to use Aspose.Cells Workbook and Worksheet APIs to apply a predefined chart template to several Excel files in one run, specifying the data range and chart placement programmatically.
// Common Searches: how to apply a .crtx chart template to multiple Excel files using Aspose.Cells C# | batch insert chart template into several workbooks with Aspose.Cells .NET | C# loop through Excel workbooks and add chart from template file | Aspose.Cells add chart from .crtx to first worksheet of each workbook
// Tags: Aspose.Cells batch chart template insertion | chart template byte array loading C# | Charts.Add method with template Aspose.Cells | apply chart template to multiple .xlsx files | set custom chart title Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // This C# program reads a .crtx chart template into a byte array, loops through an array of .xlsx workbook paths, adds the template as a chart to the first worksheet of each workbook with a defined data range and position, sets a custom chart title, and saves the updated workbooks.
class AddChartTemplateToWorkbooks
{
    static void Main()
    {
        // Paths to the workbooks that need the chart template applied
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        const string templatePath = "ChartTemplate.crtx";

        // Verify that the chart template file exists
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // Load the chart template file (.crtx) into a byte array
        byte[] chartTemplateData;
        try
        {
            chartTemplateData = File.ReadAllBytes(templatePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading template file: {ex.Message}");
            return;
        }

        // Define the data range that the chart will use (adjust as needed)
        const string dataRange = "A1:B5";

        // Define the position of the chart within the worksheet
        const int topRow = 5;
        const int leftColumn = 0;
        const int bottomRow = 20;
        const int rightColumn = 8;

        foreach (string filePath in workbookFiles)
        {
            // Verify that the workbook file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Workbook file not found, skipping: {filePath}");
                continue;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(filePath);

                // Use the first worksheet (or modify to select a specific one)
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a chart using the preset template
                int chartIndex = worksheet.Charts.Add(
                    chartTemplateData,   // Template byte array
                    dataRange,           // Data range for the chart
                    true,                // Plot series by column (vertical)
                    topRow, leftColumn, bottomRow, rightColumn);

                // Optional: customize the newly added chart
                Chart chart = worksheet.Charts[chartIndex];
                chart.Title.Text = "Chart from Template";

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath);
                Console.WriteLine($"Chart added and workbook saved: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook '{filePath}': {ex.Message}");
            }
        }
    }
}
