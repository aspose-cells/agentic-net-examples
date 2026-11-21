// Title: Map worksheet names to their SheetId indices with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get worksheet index by name | How to create a dictionary of worksheet names and sheet IDs in Aspose.Cells | Retrieve SheetId for a specific worksheet using Aspose.Cells .NET | Iterate through worksheets and store their indices in a C# dictionary | Case‑insensitive lookup of worksheet SheetId in an Aspose.Cells workbook | Aspose.Cells example for mapping worksheet names to sheet indexes
// Tags: worksheet name to sheetid mapping Aspose.Cells | c# dictionary of worksheet indices | iterate workbook worksheets Aspose.Cells | case-insensitive worksheet lookup .NET | aspocells sheet index retrieval example

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetIdMapperApp
{
    // The sample loads an Excel file with Aspose.Cells, iterates over all worksheets, and builds a case‑insensitive Dictionary<string,int> where each key is the worksheet name and each value is its SheetId (index). It then prints the mapping and includes basic file‑existence checking and exception handling.
    class WorksheetIdMapper
    {
        static void Main()
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Dictionary to hold worksheet name as key and its unique index as value
                Dictionary<string, int> sheetIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Iterate through all worksheets and store their names and indices
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheetIdMap[sheet.Name] = sheet.Index;
                }

                // Output the mapping
                foreach (KeyValuePair<string, int> kvp in sheetIdMap)
                {
                    Console.WriteLine($"Worksheet: {kvp.Key}, SheetId: {kvp.Value}");
                }

                // (Optional) Save the workbook if any changes were made
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
