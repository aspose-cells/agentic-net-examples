// Title: Load an Excel workbook and remove all conditional formatting rules using Aspose.Cells for .NET to improve load performance
// Common Searches: Aspose.Cells C# remove all conditional formatting from a loaded workbook | how to clear conditional formatting in each worksheet with Aspose.Cells .NET | speed up Excel file processing by stripping conditional formats using Aspose.Cells | C# code to load an .xlsx and delete all conditional formatting rules | Aspose.Cells example for clearing ConditionalFormattings collection
// Tags: Aspose.Cells clear ConditionalFormattings collection | remove conditional formatting from Excel workbook .NET | optimize workbook load speed by deleting conditional formats | C# Aspose.Cells conditional formatting removal | strip Excel conditional formats using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file (or creates a new workbook if the file is missing), iterates through each worksheet, clears all conditional formatting rules via the ConditionalFormattings.Clear() method, and saves the modified workbook. It includes handling for Aspose.Cells‑specific exceptions and generic runtime errors, helping improve processing speed for analysis pipelines.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Remove all conditional formatting rules from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ConditionalFormattings.Clear();
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., password protected workbook)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle all other runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
