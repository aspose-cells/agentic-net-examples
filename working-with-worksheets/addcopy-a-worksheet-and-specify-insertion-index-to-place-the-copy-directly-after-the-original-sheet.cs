// Title: How to copy a worksheet and insert it immediately after the original sheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy worksheet and place it after the source sheet | AddCopy method to insert duplicated worksheet at a specific index in .NET | C# code to duplicate the first worksheet and rename the copy using Aspose.Cells | Save workbook after inserting a copied worksheet with Aspose.Cells | Insert copied Excel sheet directly after the original using Aspose.Cells API
// Tags: Aspose.Cells AddCopy method usage | insert duplicated worksheet at specific index | rename copied worksheet programmatically | save workbook after worksheet copy C# | duplicate worksheet with index control Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, copies the first worksheet, inserts the duplicate immediately after the original using the AddCopy method, optionally renames the new sheet, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (the one to be copied)
            Worksheet originalSheet = workbook.Worksheets[0];

            // Add a copy of the worksheet; the copy is placed directly after the source sheet
            int copiedIndex = workbook.Worksheets.AddCopy(originalSheet.Index);

            // Retrieve the newly copied worksheet
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

            // Optionally rename the copied worksheet
            copiedSheet.Name = originalSheet.Name + "_Copy";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
