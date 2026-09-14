// Title: Determine whether a worksheet is read‑only protected using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, checks the IsProtected property of a specific worksheet, and outputs the protection status. | Create a reusable method in C# that accepts a worksheet name and returns true if the worksheet is protected, using Aspose.Cells. | Modify the example to skip saving the workbook and only display the protection type for each worksheet in the workbook.
// Common Searches: aspocells c# check worksheet read only protection | how to use Aspose.Cells to detect if an Excel sheet is locked | C# Aspose.Cells IsProtected property example for .xlsx files
// Tags: Aspose.Cells IsProtected worksheet check | C# read‑only Excel worksheet detection | Excel .xlsx worksheet protection status .NET | Aspose.Cells workbook load protection query | determine worksheet protection type Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook with Aspose.Cells, examines the first worksheet's IsProtected flag, prints whether it is read‑only protected, and optionally saves the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine if the worksheet is protected
            if (sheet.IsProtected)
            {
                // Granular edit permissions are not exposed in all versions,
                // so we treat any protection as a generic protected state.
                Console.WriteLine($"Worksheet \"{sheet.Name}\" is protected. Protection type: Protected");
            }
            else
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" is not protected.");
            }

            // Save the workbook (optional if no changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
