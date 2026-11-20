// Title: C# Aspose.Cells example: Merge cells AB6‑AC8 and apply custom phone number format '(###) ###‑####'
// Common Searches: how to merge a range of cells and apply a phone number format using Aspose.Cells in C# | Aspose.Cells C# code to merge cells AB6 to AC8 and set custom number format | apply custom '(###) ###-####' format to merged cells with Aspose.Cells .NET | C# example for merging cells and formatting as phone number in an existing workbook | Aspose.Cells merge cells and assign phone number style before saving workbook
// Tags: merge cells range Aspose.Cells C# | custom phone number number format Aspose.Cells | apply style to merged range Aspose.Cells | set number format for merged cells .NET | export formatted workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook, merges the cells from AB6 to AC8 on the first worksheet, creates a style with the custom phone number format '(###) ###-####', applies this number format to the merged range using a StyleFlag, and saves the updated workbook to a new file.
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

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells AB6:AC8
            Aspose.Cells.Range mergeRange = sheet.Cells.CreateRange("AB6", "AC8");
            mergeRange.Merge();

            // Create a style with the phone number custom format
            Style phoneStyle = workbook.CreateStyle();
            phoneStyle.Custom = "(###) ###-####";

            // Apply the style only to the number format of the merged range
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            mergeRange.ApplyStyle(phoneStyle, flag);

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
