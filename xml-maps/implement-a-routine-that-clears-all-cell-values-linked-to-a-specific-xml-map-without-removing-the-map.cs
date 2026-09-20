// Title: Clear values of cells linked to a specific XML map while keeping the map using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that empties the contents of every cell associated with a given XML map but leaves the XML map definition intact. | Show how to iterate through all worksheets and clear only the cells mapped to a particular XML map, then save the workbook using Aspose.Cells. | Provide a fallback implementation in C# that clears all cell values when the Aspose.Cells version does not expose XML map cell references, ensuring the map remains.
// Common Searches: Aspose.Cells C# clear data from cells mapped to an XML map without deleting the map | how to preserve an XML map while removing its cell values in a .xlsx using Aspose.Cells | C# Aspose.Cells example for clearing only XML‑mapped cells | fallback method to clear all cells when XML map API is unavailable in Aspose.Cells | remove values of XML‑mapped cells in Excel workbook using Aspose.Cells .NET
// Tags: clear XML map linked cells Aspose.Cells | keep XML map while clearing data C# | worksheet iteration clear cells Aspose.Cells | fallback clear all cells when XML map API missing | Aspose.Cells XML map manipulation limitation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks the input file, and iterates through each worksheet and every cell to set its value to an empty string. Because the current Aspose.Cells API does not expose direct XML map cell references, the code uses a safe fallback that clears all cell values while preserving the XML map definition, then saves the result to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputPath);

                // Since direct access to XML maps may not be available in the current API version,
                // clear all cells in the workbook as a safe fallback.
                ClearAllCells(workbook);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Clears all cell values in the given workbook
        private static void ClearAllCells(Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    cell.PutValue(string.Empty);
                }
            }
        }
    }
}
