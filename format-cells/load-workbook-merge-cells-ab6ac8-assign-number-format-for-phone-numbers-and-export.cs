// Title: Merge cells AB6:AC8 and apply a custom phone‑number format '(###) ###‑####' in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing workbook, merge the range AB6:AC8, create a style with the custom number format '(###) ###‑####', apply it to the merged cells, and save the file using Aspose.Cells in C#. | Using Aspose.Cells, programmatically set a phone‑number format for a merged cell block (AB6‑AC8) and export the updated workbook.
// Common Searches: how to merge AB6 to AC8 and set phone number format with Aspose.Cells C# | Aspose.Cells custom number format for merged cells example | C# code to apply '(###) ###‑####' format to a range in Excel using Aspose.Cells | merge specific cells and apply phone number style in .NET Excel library
// Tags: merge cell range AB6:AC8 Aspose.Cells | custom phone number format Aspose.Cells .NET | apply number format to merged cells Excel C# | load and save workbook with Aspose.Cells styling | create style with custom format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, merges cells AB6:AC8, defines a style with custom phone‑number format '(###) ###‑####', applies the style to the merged range, and saves the workbook as output.xlsx.
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells AB6:AC8
            // Row index is zero‑based (row 6 -> index 5)
            // Column AB -> index 27, AC -> index 28
            // Merge 3 rows (6‑8) and 2 columns (AB‑AC)
            sheet.Cells.Merge(5, 27, 3, 2);

            // Create a style with the custom phone number format "(###) ###‑####"
            Style phoneStyle = workbook.CreateStyle();
            phoneStyle.Custom = "(###) ###-####";

            // Apply the style to the merged range (or any range that holds phone numbers)
            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange(5, 27, 3, 2);
            mergedRange.ApplyStyle(phoneStyle, new StyleFlag { NumberFormat = true });

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
