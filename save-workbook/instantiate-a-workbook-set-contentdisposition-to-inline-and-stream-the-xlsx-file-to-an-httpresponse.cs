// Title: Create an Aspose.Cells workbook, write text to cell A1, and save it as an XLSX file to a specified directory in C#
// Common Searches: C# Aspose.Cells how to create a new workbook and write to cell A1 | saving an Aspose.Cells workbook to a custom output folder in C# | ensure target directory exists before saving an XLSX file with Aspose.Cells | wrap errors when exporting a workbook using Aspose.Cells in C# | Aspose.Cells example for writing data and saving to file path
// Tags: Aspose.Cells create workbook C# | Aspose.Cells write cell A1 C# | Aspose.Cells save workbook as XLSX | C# create output directory before file save | Aspose.Cells export error handling C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to instantiate an Aspose.Cells Workbook in C#, write "Hello Aspose.Cells" into cell A1 of the first worksheet, ensure the destination folder exists, and save the workbook as an XLSX file to a given path while handling exceptions.
public class WorkbookExporter
{
    /// <param name="outputPath">Full path where the workbook will be saved.</param>
    public void ExportWorkbook(string outputPath)
    {
        try
        {
            // Ensure the directory exists.
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Create a new workbook and add sample data.
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells");

            // Save the workbook directly to the file system in XLSX format.
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to provide context while preserving the original exception.
            throw new InvalidOperationException($"Failed to export workbook to '{outputPath}'.", ex);
        }
    }
}

public static class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Define output path (creates an Output folder in the current directory).
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Output", "SampleWorkbook.xlsx");
            var exporter = new WorkbookExporter();
            exporter.ExportWorkbook(outputPath);
            Console.WriteLine($"Workbook successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
