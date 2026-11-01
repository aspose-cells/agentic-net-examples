// Title: How to rename an existing named range and update all formula references using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# rename named range and refresh formulas | change named range name in Excel workbook with Aspose.Cells .NET | update all formula references after renaming a named range using Aspose.Cells | C# code to rename Excel named range and recalculate workbook | Aspose.Cells rename Name object and recalculate formulas
// Tags: rename named range Aspose.Cells .NET | recalculate workbook formulas Aspose.Cells | update named range references C# | Aspose.Cells Name.Text property usage | save workbook after range rename Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, retrieves the named range "OldName", renames it to "NewName" via the Name.Text property, recalculates all formulas to reflect the new name, ensures the output directory exists, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Define input and output file paths (replace placeholders with actual paths)
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range with the old name
            Name oldRange = workbook.Worksheets.Names["OldName"];
            if (oldRange != null)
            {
                // Rename the range by setting its Text property
                oldRange.Text = "NewName";
            }

            // Recalculate formulas so that any references to the renamed range are updated
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
