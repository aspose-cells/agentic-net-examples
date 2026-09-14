// Title: Copy a range with formulas to another workbook and automatically update external workbook references using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to copy the A1:C10 range from source.xlsx to dest.xlsx while preserving all formulas. | Iterate over the copied cells in dest.xlsx and replace any "[source.xlsx]" workbook name in the formulas with the destination file name. | Save the updated destination workbook as dest.xlsx.
// Common Searches: how to copy a range with formulas to a new workbook using Aspose.Cells C# | replace source workbook name in cell formulas after copying range Aspose.Cells .NET | update external references in copied formulas Aspose.Cells example | copy range preserving formulas and change file reference Aspose.Cells C#
// Tags: range.Copy preserving formulas Aspose.Cells | external reference replacement in cell formulas .NET | copy range between workbooks Aspose.Cells C# | modify workbook name in formulas programmatically | Aspose.Cells update formula references after copy

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads source.xlsx, creates an empty destination workbook, copies the A1:C10 range while keeping formulas, iterates the copied cells to replace "[source.xlsx]" with the destination file name in each formula, and saves the result as dest.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "dest.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook that contains the range with formulas
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Ensure the destination workbook has at least one worksheet
            Worksheet destSheet = destinationWorkbook.Worksheets[0];

            // Define the range to copy from the source workbook (adjust the address as needed)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:C10");

            // Create a matching range in the destination worksheet
            AsposeRange destRange = destSheet.Cells.CreateRange("A1:C10");

            // Copy the range, preserving formulas
            sourceRange.Copy(destRange);

            // Update external references inside the copied formulas manually
            foreach (Cell cell in destRange)
            {
                if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                {
                    // Replace the source workbook name with the destination workbook name in the formula
                    string updatedFormula = cell.Formula.Replace("[source.xlsx]", $"[{Path.GetFileName(destPath)}]");
                    cell.Formula = updatedFormula;
                }
            }

            // Save the destination workbook
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to \"{destPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
