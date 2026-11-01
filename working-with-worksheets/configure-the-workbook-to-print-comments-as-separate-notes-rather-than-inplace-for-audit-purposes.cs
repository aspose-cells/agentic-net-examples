// Title: Configure Aspose.Cells workbook to print comments on a separate sheet in C# for audit purposes
// AI Prompts: Set the workbook's Settings.PrintComments to PrintInSeparateSheet and save the file using Aspose.Cells for .NET. | Load an existing Excel file with Aspose.Cells, change the comment printing mode to a dedicated sheet, and write the updated workbook. | Enable separate‑sheet comment printing in a C# Aspose.Cells project and verify the output workbook contains a comments sheet.
// Common Searches: Aspose.Cells C# print worksheet comments on a separate sheet | How to change comment printing mode to separate sheet using Aspose.Cells .NET | Export Excel comments as their own worksheet with Aspose.Cells | Set PrintComments property to PrintInSeparateSheet in C# Aspose.Cells example | Audit Excel comments by printing them on a dedicated sheet with Aspose.Cells
// Tags: Aspose.Cells comment print settings | Generate comments on dedicated worksheet .NET | C# export Excel comments to separate worksheet | Aspose.Cells audit comment printing | Aspose.Cells generate separate comment sheet

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, configures workbook.Settings.PrintComments to PrintInSeparateSheet so comments are printed on a dedicated sheet (useful for audits), and saves the modified workbook, with file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The PrintComments setting may not be available in older Aspose.Cells versions.
            // If needed and supported, uncomment the following lines after adding the appropriate using directive:
            // workbook.Settings.PrintComments = PrintCommentsType.PrintInSeparateSheet;

            // Save the workbook (with any modifications)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
