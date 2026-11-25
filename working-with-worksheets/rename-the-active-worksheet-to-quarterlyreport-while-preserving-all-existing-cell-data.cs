// Title: Rename the active worksheet to "QuarterlyReport" in an Excel workbook using Aspose.Cells for .NET without losing cell data
// Common Searches: C# Aspose.Cells rename active worksheet to specific name | How to change worksheet name while keeping data in Aspose.Cells | Set active sheet name Aspose.Cells .NET example | Preserve cell contents when renaming Excel sheet with Aspose.Cells | Rename active worksheet programmatically using Aspose.Cells for .NET
// Tags: active sheet name change Aspose.Cells | worksheet name assignment C# Aspose.Cells | preserve cell data during sheet rename | ActiveSheetIndex usage Aspose.Cells | rename Excel sheet without data loss Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook, obtains the active worksheet via ActiveSheetIndex, sets its Name property to "QuarterlyReport", and saves the file, ensuring all existing cell data remains intact.
class Program
{
    static void Main(string[] args)
    {
        // Expect two arguments: input Excel file path and output Excel file path
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <input.xlsx> <output.xlsx>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            // Load the workbook from the specified input file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the index of the active worksheet
            int activeSheetIndex = workbook.Worksheets.ActiveSheetIndex;

            // Access the active worksheet
            Worksheet activeSheet = workbook.Worksheets[activeSheetIndex];

            // Rename the active worksheet while preserving all cell data
            activeSheet.Name = "QuarterlyReport";

            // Save the modified workbook to the specified output file
            workbook.Save(outputPath);

            Console.WriteLine($"Successfully renamed the active worksheet to 'QuarterlyReport' and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
