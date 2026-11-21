// Title: Batch add a column chart to multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add column chart to each workbook in a folder | How to programmatically insert the same chart into many .xlsx files using Aspose.Cells | Iterate over Excel files and apply a chart template with Aspose.Cells .NET | Batch processing Excel workbooks to add charts in C# | Save modified workbooks to a different directory using Aspose.Cells
// Tags: Aspose.Cells batch chart insertion C# | add column chart to Excel workbook using Aspose.Cells | process multiple .xlsx files programmatically Aspose.Cells | save modified workbooks to separate folder Aspose.Cells | apply chart template to worksheets Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The C# program scans a directory for .xlsx files, loads each workbook with Aspose.Cells, adds a column chart to the first worksheet using a fixed data range, optionally sets a title and legend, and saves the updated workbook to a designated output folder.
class Program
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string sourceFolder = @"C:\Workbooks";
        // Optional: folder to save the modified workbooks (can be same as sourceFolder to overwrite)
        string outputFolder = @"C:\Workbooks\Modified";

        // Ensure source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the source folder
        string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Use the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Define chart placement
                int upperLeftRow = 5;
                int upperLeftColumn = 0;
                int lowerRightRow = 20;
                int lowerRightColumn = 10;

                // Add a column chart and obtain its index
                int chartIndex = sheet.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
                // Retrieve the chart object
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range for the chart (adjust as needed)
                // Here we assume data is in cells A1:B5 on the same worksheet
                chart.NSeries.Add("A1:B5", true);

                // Optional: set chart title
                chart.Title.Text = "Sample Chart";

                // Optional: set legend position
                chart.Legend.Position = LegendPositionType.Right;

                // Save the modified workbook (overwrites original or saves to new location)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("All workbooks have been processed.");
    }
}
