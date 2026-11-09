// Title: Load an Excel workbook with only its VBA project using Aspose.Cells for .NET (LoadDataFilterOptions.VBA)
// Common Searches: Aspose.Cells .NET load workbook with only VBA project using LoadDataFilterOptions | How to retrieve just the VBA macros from an .xlsx file with Aspose.Cells C# | Preserve VBA code when saving an Excel file with Aspose.Cells in C# | LoadDataFilterOptions VBA example for extracting macro project in Aspose.Cells | C# read Excel file without worksheets but keep VBA project using Aspose.Cells
// Tags: extract VBA project using LoadDataFilterOptions | preserve Excel macros when saving Aspose.Cells | read .xlsx without loading worksheets C# | access VbaProject object Aspose.Cells | load workbook with VBA project only Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Namespace containing VbaProject

// The example checks for the source file, loads the workbook with LoadDataFilterOptions.VBA so that only the VBA project is read, accesses the VbaProject via workbook.VbaProject, and saves the workbook, ensuring the macro code is retained.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_with_vba.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (VBA project is loaded automatically)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project (if further processing is required)
            VbaProject vbaProject = workbook.VbaProject;

            // Save the workbook; the VBA project is preserved in the saved file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
