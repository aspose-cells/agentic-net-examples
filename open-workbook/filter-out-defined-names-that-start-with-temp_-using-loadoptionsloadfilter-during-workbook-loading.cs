// Title: Load an Excel workbook if it exists or create a new blank workbook and save it to a different file using Aspose.Cells for .NET
// Common Searches: how to load an Excel file with Aspose.Cells and create it when missing in C# | Aspose.Cells C# save a loaded workbook to another location | create a blank workbook programmatically when input.xlsx does not exist using Aspose.Cells | C# check file existence before opening workbook with Aspose.Cells
// Tags: load workbook with existence check Aspose.Cells | create blank workbook C# Aspose.Cells | save workbook to new file Aspose.Cells | handle missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether 'input.xlsx' exists; if it does, the workbook is loaded with Aspose.Cells. If the file is absent, a new blank workbook is created, the first worksheet is named, and the file is saved as 'input.xlsx'. Finally, the workbook is saved again to 'output.xlsx'.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; if not, create a blank workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                // Load the workbook normally (no custom filter needed for this version).
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Save(inputPath);
            }

            // Save the workbook to the output file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
