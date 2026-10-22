// Title: How to rename the active worksheet to "QuarterlyReport" in an Excel file using Aspose.Cells for .NET while keeping all cell data intact
// AI Prompts: Use Aspose.Cells in C# to load a workbook, obtain the currently active sheet, assign "QuarterlyReport" to its Name property, and save the workbook without modifying any cell values. | Programmatically change the name of the active worksheet in an existing .xlsx file with Aspose.Cells, ensuring that all existing data remains unchanged after saving.
// Common Searches: C# Aspose.Cells rename the active sheet to QuarterlyReport without losing data | How to change the name of the current worksheet in a loaded workbook using Aspose.Cells .NET | Preserve cell contents while renaming a worksheet with Aspose.Cells API | Get active worksheet index and rename it in Aspose.Cells C# example
// Tags: Aspose.Cells programmatic worksheet renaming | retrieve active sheet index Aspose.Cells C# | save workbook preserving data after sheet rename | maintain cell values when renaming worksheet | worksheet Name property Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file with Aspose.Cells, retrieves the active worksheet via the ActiveSheetIndex, sets its Name to "QuarterlyReport", and saves the workbook as a new file, ensuring that all original cell data is retained.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the active worksheet using the Worksheets collection's ActiveSheetIndex
            int activeIndex = workbook.Worksheets.ActiveSheetIndex;
            Worksheet activeSheet = workbook.Worksheets[activeIndex];

            // Rename the active worksheet
            activeSheet.Name = "QuarterlyReport";

            // Save the workbook, preserving all existing cell data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
