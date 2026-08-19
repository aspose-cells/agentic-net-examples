// Title: C# – Load a One‑Dimensional String Array into Excel Cells with Aspose.Cells
// Description: This example creates a new Workbook, accesses the first Worksheet, and writes a string[] horizontally across the first row starting at cell A1 using Worksheet.Cells.ImportObjectArray (isVertical = false). The workbook is then saved as OneDimensionalStringArray.xlsx.
// Keywords: Aspose.Cells | C# | ImportObjectArray | string array to Excel | horizontal array import | load data into cells | save workbook as xlsx | Excel automation | populate first row | worksheet cells import
// Common Searches: Aspose.Cells import string array C# | how to write a one‑dimensional array to Excel with Aspose | ImportObjectArray horizontal example | load string[] into first row Aspose.Cells | C# write array to Excel cells starting at A1
// Developer Intent: Write a one‑dimensional string array horizontally into the first row of a worksheet starting at cell A1.
// Use Cases: Create a header row from a list of column names. | Export a simple report line generated in memory. | Transfer data from a CSV parser directly into Excel without looping through cells.
// AI Prompts: Generate C# code to import a one‑dimensional string array vertically into an Excel worksheet using Aspose.Cells. | Show how to import a two‑dimensional object array into a specific range with ImportObjectArray. | Explain each parameter of Worksheet.Cells.ImportObjectArray and how to handle null values in the source array.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a new Workbook, accesses the first Worksheet, and writes a string[] horizontally across the first row starting at cell A1 using Worksheet.Cells.ImportObjectArray (isVertical = false). The workbook is then saved as OneDimensionalStringArray.xlsx.
    public class LoadOneDimensionalStringArray
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare a one‑dimensional string array
                string[] stringArray = new string[] { "Alpha", "Beta", "Gamma", "Delta" };

                // Import the array into cells starting at row 0, column 0 (A1)
                // isVertical = false -> import horizontally across columns
                worksheet.Cells.ImportObjectArray(stringArray, 0, 0, false);

                // Save the workbook to a file
                string outputPath = "OneDimensionalStringArray.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadOneDimensionalStringArray.Run();
        }
    }
}
