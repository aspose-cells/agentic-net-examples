// Title: Batch Apply a .crtx Chart Template to Multiple Excel Workbooks with Aspose.Cells for .NET
// Description: Loads a .crtx chart template once, loops through a list of Excel files, inserts the template as a chart on the first worksheet of each workbook, sets a dynamic title, and saves the updated files to a target folder using Aspose.Cells for C#.
// Keywords: Aspose.Cells chart template | C# batch chart creation | apply .crtx to multiple workbooks | Aspose.Cells add chart from byte array | automate Excel chart styling | .NET Excel chart template | process multiple Excel files Aspose
// Common Searches: How to add a .crtx chart template to many Excel files with Aspose.Cells | Batch apply chart template to workbooks C# | Aspose.Cells add chart from template byte array | Save modified Excel workbooks after adding chart Aspose | Loop through Excel files and insert chart template .NET
// Developer Intent: Programmatically apply a predefined chart template to each workbook in a collection and write the modified files to an output directory.
// Use Cases: Generate consistent sales charts across monthly report workbooks. | Standardize visual style for a series of financial statements. | Automate chart formatting for a batch of dashboard Excel files.
// AI Prompts: Show how to detect the data range in each workbook and pass it to the chart template. | Provide code to add multiple charts from the same template to different worksheets within each workbook. | Explain error‑handling strategies for missing .crtx files and fallback to a default chart style.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a .crtx chart template once, loops through a list of Excel files, inserts the template as a chart on the first worksheet of each workbook, sets a dynamic title, and saves the updated files to a target folder using Aspose.Cells for C#.
class ChartTemplateApplier
{
    // Applies a chart template to each workbook in the collection and saves the result.
    public static void ApplyTemplateToWorkbooks(string[] workbookFiles, string chartTemplateFile, string outputDirectory)
    {
        // Verify that the chart template file exists.
        if (!File.Exists(chartTemplateFile))
        {
            Console.WriteLine($"Chart template file not found: {chartTemplateFile}");
            return;
        }

        byte[] templateData;
        try
        {
            // Load the chart template file once as a byte array.
            templateData = File.ReadAllBytes(chartTemplateFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read chart template file: {ex.Message}");
            return;
        }

        // Ensure the output directory exists.
        try
        {
            Directory.CreateDirectory(outputDirectory);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create output directory: {ex.Message}");
            return;
        }

        foreach (string workbookPath in workbookFiles)
        {
            // Verify that the workbook file exists.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found, skipping: {workbookPath}");
                continue;
            }

            try
            {
                // Load the workbook from file.
                using (Workbook workbook = new Workbook(workbookPath))
                {
                    // Use the first worksheet (adjust as needed).
                    Worksheet sheet = workbook.Worksheets[0];

                    // Define the data range for the chart (adjust to your data layout).
                    string dataRange = "A1:B5";

                    // Add a chart using the preset template.
                    // Parameters: template bytes, data range, isVertical, topRow, leftColumn, bottomRow, rightColumn
                    int chartIndex = sheet.Charts.Add(
                        templateData,
                        dataRange,
                        true,   // Plot series by column (vertical)
                        5,      // Top row of the chart
                        0,      // Left column of the chart
                        20,     // Bottom row of the chart
                        8       // Right column of the chart
                    );

                    // Optionally, further customize the chart after adding it.
                    Chart chart = sheet.Charts[chartIndex];
                    chart.Title.Text = Path.GetFileNameWithoutExtension(workbookPath) + " Chart";

                    // Save the modified workbook to the output folder (overwrites if same name).
                    string outputPath = Path.Combine(outputDirectory, Path.GetFileName(workbookPath));
                    workbook.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook '{workbookPath}': {ex.Message}");
            }
        }
    }

    // Example usage.
    static void Main()
    {
        // List of workbook file paths to process.
        string[] workbooks = new string[]
        {
            @"C:\Input\Report1.xlsx",
            @"C:\Input\Report2.xlsx",
            @"C:\Input\Report3.xlsx"
        };

        // Path to the chart template (.crtx) file.
        string chartTemplate = @"C:\Templates\MyChartTemplate.crtx";

        // Directory where modified workbooks will be saved.
        string outputDir = @"C:\Output";

        ApplyTemplateToWorkbooks(workbooks, chartTemplate, outputDir);

        Console.WriteLine("Chart templates applied to all processed workbooks.");
    }
}
