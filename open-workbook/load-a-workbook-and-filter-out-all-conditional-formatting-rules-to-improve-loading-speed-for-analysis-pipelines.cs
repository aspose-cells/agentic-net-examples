// Title: How to remove all conditional formatting from an Excel workbook using Aspose.Cells for .NET to boost load performance
// AI Prompts: Generate a C# program that opens an .xlsx file with Aspose.Cells, iterates through every worksheet, clears all ConditionalFormattings, and saves the cleaned workbook to a new location. | Write C# code that verifies the existence of a source workbook, creates a minimal workbook if it is missing, then uses Aspose.Cells to strip every conditional formatting rule to reduce file size and improve subsequent loading speed.
// Common Searches: asp.net remove all conditional formatting from excel workbook using Aspose.Cells | c# strip conditional formatting from each sheet to improve workbook loading speed | how to clear conditional formatting with Aspose.Cells before saving | optimize excel file load time by deleting conditional formats in C# Aspose.Cells
// Tags: Aspose.Cells clear ConditionalFormattings | remove conditional formatting .xlsx C# | optimize workbook load speed Aspose.Cells | strip Excel conditional formats .NET | conditional formatting cleanup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a source .xlsx file, creates a minimal workbook if needed, loads the workbook with Aspose.Cells, removes all conditional formatting from every worksheet, and saves the cleaned file to improve loading performance.
class Program
{
    static void Main()
    {
        // Paths for source and cleaned workbooks
        string inputPath = @"C:\Data\SourceWorkbook.xlsx";
        string outputPath = @"C:\Data\CleanedWorkbook.xlsx";

        try
        {
            // Ensure the source file exists; create a minimal workbook if it does not
            if (!File.Exists(inputPath))
            {
                // Create directory if needed
                Directory.CreateDirectory(Path.GetDirectoryName(inputPath));

                // Create a new workbook with a default worksheet and save it as the source file
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(inputPath);
            }

            // Load the workbook from the input path
            Workbook workbook = new Workbook(inputPath);

            // Remove all conditional formatting from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ConditionalFormattings.Clear();
            }

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook cleaned and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
