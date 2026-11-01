// Title: Assign Excel worksheet TabId values chronologically based on custom CreatedDate property or workbook creation time with Aspose.Cells for .NET
// Common Searches: how to assign TabId to worksheets based on a custom CreatedDate property in Aspose.Cells C# | sorting Excel sheets by creation date and updating TabId using Aspose.Cells .NET | fallback to workbook BuiltInDocumentProperties.CreatedTime when worksheet custom property is missing Aspose.Cells | programmatically reorder worksheet tabs chronologically in a .xlsx file with Aspose.Cells
// Tags: set worksheet TabId by creation date Aspose.Cells | sort worksheets using custom property CreatedDate Aspose.Cells | fallback to BuiltInDocumentProperties.CreatedTime Aspose.Cells | chronological Excel tab ordering .NET | update Excel sheet TabId programmatically Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// The example loads an Excel workbook with Aspose.Cells, reads each worksheet's creation date from a custom "CreatedDate" property (or falls back to the workbook's built‑in creation time), sorts the worksheets chronologically, assigns sequential 1‑based TabId values according to that order, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Collect each worksheet with its associated creation date
            var sheetInfos = new List<(Worksheet sheet, DateTime created)>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                DateTime createdDate;

                // Attempt to read a custom property named "CreatedDate" from the worksheet
                CustomProperty prop = null;
                try
                {
                    prop = sheet.CustomProperties["CreatedDate"];
                }
                catch
                {
                    // Property not found; prop remains null
                }

                if (prop != null && prop.Value != null)
                {
                    // Use the value of the custom property
                    createdDate = Convert.ToDateTime(prop.Value);
                }
                else
                {
                    // Fallback to the workbook's built‑in creation time
                    createdDate = workbook.BuiltInDocumentProperties.CreatedTime;
                }

                sheetInfos.Add((sheet, createdDate));
            }

            // Sort worksheets chronologically by their creation dates
            var orderedSheets = sheetInfos.OrderBy(info => info.created).ToList();

            // Update TabId sequentially according to the chronological order (1‑based)
            for (int i = 0; i < orderedSheets.Count; i++)
            {
                orderedSheets[i].sheet.TabId = i + 1;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
