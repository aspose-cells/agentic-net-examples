// Title: Convert a cell range to an Excel ListObject (table) with Aspose.Cells for .NET to keep charts dynamic
// AI Prompts: Write C# code using Aspose.Cells that loads an existing .xlsx, defines a specific range, adds a ListObject with a display name and TableStyleMedium9, and saves the workbook. | Show how to create a named Excel table from A1:D10 in a worksheet and apply a built‑in table style using Aspose.Cells. | Provide a try‑catch example that checks the input file, creates a ListObject, ensures the output folder exists, and writes the updated file.
// Common Searches: Aspose.Cells how to add an Excel table from a range for dynamic chart source | C# create ListObject A1:D10 and apply built‑in style with Aspose.Cells | convert data range to table Aspose.Cells .NET example | save workbook after inserting ListObject using Aspose.Cells | ensure output directory exists before saving workbook Aspose.Cells C#
// Tags: Aspose.Cells add ListObject from range | Aspose.Cells apply built‑in table style | Aspose.Cells dynamic chart source table | C# verify input file before loading workbook | C# create output directory for Aspose.Cells save

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an existing XLSX file, defines the range A1:D10, adds a ListObject named "DataTable" with TableStyleMedium9, ensures the output directory exists, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the table (A1:D10)
            int firstRow = 0;      // Row 1 (zero‑based)
            int firstColumn = 0;   // Column A (zero‑based)
            int lastRow = 9;       // Row 10 (zero‑based)
            int lastColumn = 3;    // Column D (zero‑based)

            // Add a ListObject (Excel Table) covering the specified range
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a display name to the table (used as the table name in Excel)
            table.DisplayName = "DataTable";

            // Apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
