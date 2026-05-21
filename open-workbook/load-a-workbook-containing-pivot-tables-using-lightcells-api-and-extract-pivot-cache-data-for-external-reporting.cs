using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ExtractPivotCache
{
    static void Main()
    {
        const string inputFile = "PivotData.xlsx";
        const string outputFile = "PivotData_Extracted.xlsx";

        // Verify that the input workbook exists.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load options: enable parsing of pivot cached records.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                ParsingPivotCachedRecords = true
            };

            // Load the workbook.
            var workbook = new Workbook(inputFile, loadOptions);

            // (Optional) Iterate through worksheets and pivot tables if further processing is needed.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Example: access pivot table name (no modification here).
                    Console.WriteLine($"Found pivot table \"{pivot.Name}\" on sheet \"{sheet.Name}\".");
                }
            }

            // Save the workbook with extracted pivot cache data.
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}