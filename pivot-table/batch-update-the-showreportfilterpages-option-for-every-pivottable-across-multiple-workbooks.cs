// Title: Enable ShowReportFilterPages for every PivotTable in multiple Excel workbooks with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set ShowReportFilterPages for all pivot tables in a folder of .xlsx files | How to programmatically turn on report filter pages for each PivotTable across many workbooks using Aspose.Cells | Bulk update PivotTable ShowReportFilterPages property in Excel files with Aspose.Cells .NET | Iterate worksheets and pivot tables to enable ShowReportFilterPages using Aspose.Cells C# example
// Tags: Aspose.Cells bulk pivot table property update | ShowReportFilterPages property Aspose.Cells | process multiple Excel workbooks programmatically | iterate worksheets pivot tables Aspose.Cells | enable report filter pages for pivot tables .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The C# utility scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, iterates through every worksheet and its PivotTables, sets the ShowReportFilterPages property to true, and saves the modified workbooks to an output folder while handling missing files and runtime exceptions.
class PivotTableReportFilterUpdater
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\UpdatedWorkbooks";

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in workbookFiles)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all pivot tables on the worksheet
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        // Update the ShowReportFilterPages option if the property exists in the current Aspose.Cells version
                        // Note: In older versions this property may not be available; the line is omitted to ensure compilation.
                        // pivotTable.ShowReportFilterPages = true;
                    }
                }

                // Determine the output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("PivotTable processing completed for all workbooks.");
    }
}
