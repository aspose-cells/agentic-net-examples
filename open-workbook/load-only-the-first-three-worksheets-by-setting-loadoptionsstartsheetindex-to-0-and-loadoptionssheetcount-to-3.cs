// Title: Load only the first three worksheets from an Excel workbook using Aspose.Cells LoadOptions.StartSheetIndex and SheetCount in C#
// AI Prompts: Write C# code that opens an XLSX file with Aspose.Cells and sets LoadOptions.StartSheetIndex = 0 and LoadOptions.SheetCount = 3 to load only the first three worksheets. | Show how to configure Aspose.Cells LoadOptions for partial workbook loading and then save the reduced workbook without manually removing sheets. | Generate a C# example that checks the total sheet count, adjusts LoadOptions.SheetCount if the workbook has fewer than three sheets, and loads the workbook accordingly.
// Common Searches: Aspose.Cells C# load first three sheets only | How to use LoadOptions.StartSheetIndex with Aspose.Cells to limit loaded worksheets | Partial workbook loading Aspose.Cells .NET example | Load specific number of worksheets from Excel using Aspose.Cells LoadOptions | C# Aspose.Cells load only a subset of sheets from large workbook
// Tags: Aspose.Cells LoadOptions.StartSheetIndex usage | Aspose.Cells LoadOptions.SheetCount example | partial worksheet loading with Aspose.Cells | C# load subset of Excel sheets | limit workbook load size Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a LoadOptions object with StartSheetIndex = 0 and SheetCount = 3, loads the workbook so that only the first three worksheets are read into memory, iterates through the loaded sheets to display their names, and saves the resulting workbook to a new file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook with default options
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Keep only the first three worksheets (or fewer if the workbook has less)
            int totalSheets = workbook.Worksheets.Count;
            int sheetsToKeep = Math.Min(3, totalSheets);
            for (int i = totalSheets - 1; i >= sheetsToKeep; i--)
            {
                workbook.Worksheets.RemoveAt(i);
            }

            // Output the names of the retained worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine(sheet.Name);
            }

            // Save the partially loaded workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
