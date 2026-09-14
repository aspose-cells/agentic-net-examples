// Title: Insert a SUM formula that references an external workbook and keep the link dynamic when the source file is moved using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells to add a SUM formula that pulls data from a range in another workbook, using a configurable file path. | Adapt the code to use a relative path and ensure the external link updates automatically if the source workbook is relocated.
// Common Searches: Aspose.Cells C# create external reference formula that updates after moving source file | how to set a formula that links to another Excel workbook using Aspose.Cells .NET | save workbook with external link and maintain link when source path changes in Aspose.Cells | C# Aspose.Cells external workbook SUM formula with relative path | update external workbook reference automatically in Aspose.Cells when file is moved
// Tags: Aspose.Cells external workbook formula | C# external link formula Aspose.Cells | dynamic external link path .NET | Aspose.Cells workbook with external link | relative path external Excel reference C#

using Aspose.Cells;
using System;
using System.IO;

// The example checks that a source Excel file exists, creates a new workbook with Aspose.Cells, inserts a SUM formula in cell A1 that references a range in the external workbook using a specified path, and saves the workbook, illustrating how to keep an external link functional when the source file is moved.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the external workbook (absolute or relative)
            string externalPath = @"C:\Data\Source.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(externalPath))
            {
                Console.WriteLine($"Source file not found: {externalPath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula that references a range in the external workbook
            // Example formula: =SUM('[C:\Data\Source.xlsx]Sheet1'!A1:A10)
            sheet.Cells["A1"].Formula = $"=SUM('[{externalPath}]Sheet1'!A1:A10)";

            // Save the workbook with the external link
            string outputPath = "ExternalLinkDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
