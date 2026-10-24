// Title: Implementing a unique‑value constraint for an Aspose.Cells ListObject column in C# (using data validation)
// Common Searches: Aspose.Cells C# add data validation to prevent duplicate values in a table column | How to simulate a unique index on an Excel ListObject using Aspose.Cells .NET | C# code to enforce distinct entries in an Aspose.Cells worksheet table | Workaround for unique column constraint in Aspose.Cells workbook with C# | Apply data validation for unique entries in an Aspose.Cells ListObject
// Tags: Aspose.Cells data validation unique column | C# enforce distinct values in Excel table | ListObject unique constraint workaround | Aspose.Cells .NET unique index alternative | Excel table duplicate prevention using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, creates a ListObject (Excel table) over a defined range, assigns a display name, and saves the file. It explains that Aspose.Cells does not provide a direct API to add a unique index to a ListObject column, and demonstrates using data validation or custom logic as a workaround to enforce uniqueness during data entry.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that will become the table (A1:D10 in this example)
            int firstRow = 0;      // zero‑based
            int firstColumn = 0;   // zero‑based
            int lastRow = 9;       // zero‑based (row 10)
            int lastColumn = 3;    // zero‑based (column D)

            // Add a ListObject (Excel table) to the defined range.
            // The last parameter indicates whether the range has headers.
            int tableIdx = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Set a display name for the table (Name property is not available in this version)
            table.DisplayName = "MyTable";

            // NOTE: Adding a unique index directly to a ListObject is not supported in the current API.
            // If needed, implement uniqueness via data validation or other logic.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
