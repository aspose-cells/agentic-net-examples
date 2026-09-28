// Title: How to delete the 'SummaryData' named range from an Excel workbook and confirm its removal using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that removes the named range 'SummaryData' from a workbook and checks that it no longer exists. | Generate a C# snippet that loads an .xlsx file, deletes a specific named range, validates the deletion, and saves the updated file using Aspose.Cells.
// Common Searches: how to delete a named range using Aspose.Cells in C# | verify that a named range has been removed from an Excel workbook with Aspose.Cells | C# example for removing a particular named range from an .xlsx file via Aspose.Cells
// Tags: Aspose.Cells named range removal C# | programmatic Excel named range deletion | validate named range absence Aspose.Cells | workbook Names collection update Aspose.Cells | C# Aspose.Cells delete specific range .xlsx

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, checks for the named range "SummaryData", removes it if present, confirms the range is no longer in the workbook's Names collection, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove the named range "SummaryData" if it exists
            if (workbook.Worksheets.Names["SummaryData"] != null)
            {
                workbook.Worksheets.Names.Remove("SummaryData");
            }

            // Verify that the named range no longer appears in the collection
            bool exists = workbook.Worksheets.Names["SummaryData"] != null;
            Console.WriteLine("SummaryData exists after removal: " + exists); // should output False

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
