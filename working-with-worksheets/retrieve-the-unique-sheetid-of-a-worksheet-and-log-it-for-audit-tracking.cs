// Title: Retrieve a worksheet's unique index (SheetId) and log it for audit tracking with Aspose.Cells for .NET
// Common Searches: how to obtain worksheet index as identifier with Aspose.Cells in C# | log Excel worksheet unique id for audit using Aspose.Cells .NET | retrieve worksheet.Index value for tracking changes in Aspose.Cells | C# Aspose.Cells get worksheet unique identifier for logging | audit trail worksheet id Aspose.Cells workbook
// Tags: Aspose.Cells worksheet index extraction | C# audit log worksheet unique id | Worksheet.Index as SheetId in .NET | track Excel sheet identifier with Aspose.Cells | unique worksheet identifier for logging

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example verifies the presence of an input Excel file, loads it into an Aspose.Cells Workbook, accesses the first worksheet, uses the worksheet's zero‑based Index property as a unique SheetId, and writes the worksheet name and identifier to the console. It includes exception handling to ensure robust audit‑ready logging.
    class Program
    {
        static void Main()
        {
            const string filePath = "input.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Input file not found: {filePath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Aspose.Cells for .NET does not expose a SheetId property.
                // Use the worksheet's zero‑based Index as a unique identifier.
                int sheetId = worksheet.Index;

                // Log the identifier for audit tracking
                Console.WriteLine($"Worksheet '{worksheet.Name}' has SheetId (Index): {sheetId}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
