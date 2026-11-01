// Title: Set 120% zoom for every worksheet and export each sheet as an individual XLS file with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set worksheet zoom to 120 percent and save each sheet as its own XLS file | How to split an Excel workbook into separate XLS files while applying a custom zoom level using Aspose.Cells | Aspose.Cells copy a single worksheet to a new workbook and change the zoom factor in C# | Programmatically export each worksheet to separate .xls files with a specific zoom setting using Aspose.Cells | Apply 120% zoom to all sheets in a workbook before saving them individually with Aspose.Cells .NET
// Tags: worksheet zoom setting Aspose.Cells C# | export single worksheet to XLS Aspose.Cells | split workbook into separate files Aspose.Cells | apply custom zoom factor to sheets Aspose.Cells | copy worksheet to new workbook C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, iterates through all worksheets, sets each sheet's Zoom property to 120%, creates a new workbook containing only that sheet, and saves it as a separate .xls file named after the worksheet, with error handling for missing files and per‑sheet exceptions.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the source workbook
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                Worksheet sheet = sourceWorkbook.Worksheets[i];

                try
                {
                    // Apply 120% zoom level to the current worksheet
                    sheet.Zoom = 120;

                    // Create a new workbook to hold only this worksheet
                    Workbook singleSheetWorkbook = new Workbook();

                    // Remove the default empty sheet that Aspose.Cells creates
                    singleSheetWorkbook.Worksheets.Clear();

                    // Copy the current worksheet into the new workbook by name
                    singleSheetWorkbook.Worksheets.AddCopy(sheet.Name);

                    // Define the output file name (one XLS file per worksheet)
                    string outputFileName = $"{sheet.Name}.xls";

                    // Save the new workbook; format is inferred from the file extension
                    singleSheetWorkbook.Save(outputFileName);
                    Console.WriteLine($"Saved worksheet '{sheet.Name}' to '{outputFileName}'.");
                }
                catch (Exception ex)
                {
                    // Handle errors for individual worksheets
                    Console.WriteLine($"Error processing worksheet '{sheet.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
