// Title: C# – Access a Worksheet by Zero‑Based Index with Aspose.Cells for .NET
// Description: This example shows how to create a new Workbook, retrieve the first Worksheet using its zero‑based index (0), write a text value to cell A1, and save the file. It demonstrates the simplest way to access and manipulate a worksheet by index using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | access worksheet by index | zero based index | retrieve worksheet | set cell value | save workbook | worksheet collection | Aspose.Cells example
// Common Searches: Aspose.Cells get worksheet by index C# | How to access first worksheet in Aspose.Cells | Write to a cell after retrieving worksheet by index | Save workbook after modifying worksheet Aspose.Cells | C# Aspose.Cells example for worksheet collection
// Developer Intent: Retrieve a worksheet from a workbook by its zero‑based index and modify a cell.
// Use Cases: Create a new workbook and write a greeting to the default sheet (index 0). | Iterate through all worksheets by index to apply uniform formatting or data entry. | Open an existing workbook, select a specific sheet by index, populate data, and export the result.
// AI Prompts: Generate C# code that accesses the third worksheet (index 2) in an Aspose.Cells workbook, sets cell B2 to a numeric value, and saves the file. | Provide an Aspose.Cells example that loops over every worksheet by index, writes the sheet name into cell A1 of each sheet, and saves the workbook. | Show how to safely retrieve a worksheet by index with try‑catch handling for out‑of‑range indexes in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example shows how to create a new Workbook, retrieve the first Worksheet using its zero‑based index (0), write a text value to cell A1, and save the file. It demonstrates the simplest way to access and manipulate a worksheet by index using Aspose.Cells for .NET.
    public class AccessWorksheetByIndex
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (empty workbook with a default worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet using its zero‑based index (0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Set a value in cell A1 to demonstrate access
            firstWorksheet.Cells["A1"].PutValue("Hello from the first worksheet!");

            // Save the workbook to verify the changes
            string outputPath = "AccessWorksheetByIndex.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
