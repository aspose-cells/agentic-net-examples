// Title: Delete a named worksheet from an Excel file and export the remaining workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads an .xlsx file, removes the worksheet called "SheetToRemove" if it exists, and saves the workbook as a PDF. | Show how to programmatically locate and delete a specific sheet in a workbook before calling Workbook.Save with SaveFormat.Pdf in C#.
// Common Searches: Aspose.Cells C# remove worksheet before PDF conversion | how to skip a sheet when converting Excel to PDF with Aspose.Cells | C# code to delete a sheet by name and then save workbook as PDF using Aspose.Cells | exclude specific worksheet from PDF output in Aspose.Cells .NET
// Tags: worksheet removal Aspose.Cells C# | PDF export after worksheet deletion Aspose.Cells | skip sheet during PDF conversion Aspose.Cells | C# delete Excel sheet before PDF | Aspose.Cells workbook modification prior to PDF

using System;
using System.IO;
using Aspose.Cells;

// The example loads "input.xlsx", searches for a worksheet named "SheetToRemove", removes it if found, and then saves the modified workbook as "output.pdf" in PDF format, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";
        const string sheetToRemove = "SheetToRemove";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Find the index of the worksheet to remove by name
            int sheetIndex = -1;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Name.Equals(sheetToRemove, StringComparison.OrdinalIgnoreCase))
                {
                    sheetIndex = i;
                    break;
                }
            }

            if (sheetIndex != -1)
            {
                workbook.Worksheets.RemoveAt(sheetIndex);
                Console.WriteLine($"Worksheet \"{sheetToRemove}\" removed.");
            }
            else
            {
                Console.WriteLine($"Worksheet \"{sheetToRemove}\" not found; no removal performed.");
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., Aspose.Cells errors, IO issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
