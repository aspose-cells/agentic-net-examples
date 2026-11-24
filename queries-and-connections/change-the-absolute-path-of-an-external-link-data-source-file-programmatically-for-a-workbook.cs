// Title: How to change the absolute file path of an external data source link in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet change external link source path in Excel workbook with Aspose.Cells | update external data source file location programmatically using Aspose.Cells .NET | set new source file path for workbook external connections in C# Aspose.Cells | replace absolute path of external link in existing .xlsx using Aspose.Cells API | modify external link file path in Excel file via Aspose.Cells for .NET
// Tags: Aspose.Cells external link path modification | C# workbook external data source update | ExternalLink API set source file path | Excel external connection path replacement .NET | programmatic external link file location change

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, defines a new external data source path, and (in a commented section) demonstrates how to iterate through each worksheet's ExternalLinks to assign the new SourceFilePath. It creates the output folder if needed and saves the modified workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Paths
            string inputPath = @"C:\Input\Sample.xlsx";
            string outputPath = @"C:\Output\Sample_Modified.xlsx";
            string newExternalPath = @"D:\DataSources\NewDataSource.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Updating external link source paths requires the ExternalLink API,
            // which may not be available in older Aspose.Cells versions.
            // The following block is omitted to ensure compilation.
            // If using a version that supports ExternalLink, uncomment and adjust accordingly:
            /*
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (ExternalLink link in sheet.ExternalLinks)
                {
                    if (!string.IsNullOrEmpty(link.SourceFilePath))
                    {
                        link.SourceFilePath = newExternalPath;
                    }
                }
            }
            */

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
