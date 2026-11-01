// Title: Copy a worksheet by its name and rename the duplicate inside the same workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing Excel file, locates a worksheet using its name, creates a copy of that worksheet in the same workbook, renames the copy, and saves the workbook with Aspose.Cells. | Write a C# snippet that verifies if a given sheet name exists; if it does not, selects the first worksheet, then duplicates it using Workbook.Worksheets.AddCopy and assigns a new name. | Provide a C# example that demonstrates copying a worksheet while keeping all cell formulas, styles, and data validations intact, then saving the updated file using Aspose.Cells.
// Common Searches: Aspose.Cells C# copy worksheet by name and rename the copy | How to duplicate a sheet inside the same workbook using Aspose.Cells for .NET | C# code to fallback to first worksheet when sheet name is missing in Aspose.Cells | Maintain cell formulas and styles during worksheet duplication using Aspose.Cells C# | Save workbook after adding a copied worksheet with Aspose.Cells C#
// Tags: worksheet copy addcopy Aspose.Cells | rename duplicated worksheet C# | fallback to first sheet Aspose.Cells | preserve formulas styles during sheet copy | save workbook after sheet duplication Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel file (or creates one if missing), retrieves a worksheet named "SourceSheet" (or the first sheet if the name is absent), copies it within the same workbook using AddCopy, renames the duplicate by appending "_Copy", and saves the result to "output.xlsx" while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string sourceSheetName = "SourceSheet";

            // Ensure the input file exists; if not, create a simple workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = sourceSheetName; // create a default sheet
                workbook.Save(inputPath);
            }

            // Retrieve the source worksheet; fall back to the first sheet if not found.
            Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];
            if (sourceSheet == null)
            {
                sourceSheet = workbook.Worksheets[0];
                Console.WriteLine($"Worksheet \"{sourceSheetName}\" not found. Using first worksheet \"{sourceSheet.Name}\" instead.");
            }

            // Copy the worksheet within the same workbook.
            int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Index);

            // Rename the copied worksheet.
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = sourceSheet.Name + "_Copy";

            // Save the workbook with the copied worksheet.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
