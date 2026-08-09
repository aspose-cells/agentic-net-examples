// Title: Batch add a .crtx chart template to Excel workbooks using Aspose.Cells for .NET (C#)
// Description: This C# utility scans a directory for .xlsx files, loads a .crtx chart template into a byte array, inserts the template as a chart on the first worksheet of each workbook, optionally customizes the title, and overwrites the original files. It streamlines the process of applying a consistent chart design across multiple Excel reports.
// Keywords: Aspose.Cells | C# | add chart template | batch process Excel files | apply .crtx chart | .crtx to workbook | automate chart insertion | iterate workbooks | Excel chart automation | chart template byte array
// Common Searches: How to add a .crtx chart template to multiple Excel files with Aspose.Cells | Batch insert predefined chart into workbooks using C# | Apply the same chart design to many .xlsx files programmatically | Aspose.Cells add chart from template to each workbook in a folder | Automate chart creation across Excel reports C#
// Developer Intent: Insert a predefined chart template into every workbook in a folder and save the updates automatically.
// Use Cases: Standardize corporate branding by embedding a company‑wide chart style into monthly sales workbooks. | Automatically enrich generated financial reports with a pre‑designed performance chart before distribution. | Refresh legacy Excel files with a new chart layout without manual editing.
// AI Prompts: Write C# code that reads a .crtx file into a byte array and adds it as a chart to all .xlsx files in a specified directory using Aspose.Cells. | Show robust error handling for batch processing Excel workbooks when inserting a chart template with Aspose.Cells for .NET. | Demonstrate how to set each inserted chart's title to the workbook filename after applying the .crtx template.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# utility scans a directory for .xlsx files, loads a .crtx chart template into a byte array, inserts the template as a chart on the first worksheet of each workbook, optionally customizes the title, and overwrites the original files. It streamlines the process of applying a consistent chart design across multiple Excel reports.
class AddChartTemplateToWorkbooks
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";

        // Path to the predefined chart template file (.crtx)
        string templatePath = @"C:\Templates\MyChartTemplate.crtx";

        // Verify that the template file exists
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // Load the chart template into a byte array
        byte[] templateData;
        try
        {
            templateData = File.ReadAllBytes(templatePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read template file: {ex.Message}");
            return;
        }

        // Verify that the input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Get all Excel files in the input folder (adjust pattern if needed)
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string workbookFile in workbookFiles)
        {
            try
            {
                // Load the workbook using the constructor that accepts a file path
                Workbook workbook = new Workbook(workbookFile);

                // Use the first worksheet (or adjust index as required)
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a chart to the worksheet using the template data.
                // Parameters:
                //   templateData   – byte[] of the .crtx file
                //   "A1:B4"        – data range for the chart (adjust to your data)
                //   true           – plot series by column (vertical)
                //   5, 0, 20, 8    – topRow, leftColumn, bottomRow, rightColumn for chart position
                int chartIndex = worksheet.Charts.Add(
                    templateData,
                    "A1:B4",
                    true,
                    5,
                    0,
                    20,
                    8);

                // Optional: further customize the chart if needed
                Chart chart = worksheet.Charts[chartIndex];
                chart.Title.Text = $"Chart added to {Path.GetFileName(workbookFile)}";

                // Save the workbook back to the same file (overwrites original)
                workbook.Save(workbookFile);
                Console.WriteLine($"Processed: {Path.GetFileName(workbookFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{workbookFile}': {ex.Message}");
                // Continue with next workbook
            }
        }

        Console.WriteLine("Chart templates have been added to all processed workbooks.");
    }
}
