// Title: Load an Excel workbook from a MemoryStream, modify a cell formula, and recalculate with custom CalculationOptions using Aspose.Cells for .NET
// AI Prompts: Read an Excel file from a MemoryStream, set cell B1 formula to '=A1*3', configure CalculationOptions with IgnoreError=true and Recursive=true, then call Workbook.CalculateFormula in C#. | Using Aspose.Cells for .NET, load a workbook from a stream, replace an existing formula, and recalculate all dependent cells with custom calculation settings.
// Common Searches: Aspose.Cells C# load Excel from MemoryStream and edit cell formula | calculate workbook formulas with ignore errors option in Aspose.Cells | how to apply recursive calculation to all worksheets using Aspose.Cells | save modified workbook to a new MemoryStream after formula recalculation Aspose.Cells
// Tags: memory stream workbook loading Aspose.Cells | modify cell formula programmatically C# | custom CalculationOptions for Workbook.CalculateFormula | ignore errors during Excel formula evaluation Aspose.Cells | recursive formula recalculation across worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Demonstrates loading an Excel workbook from a MemoryStream, updating a cell's formula, configuring CalculationOptions (IgnoreError and Recursive), and recalculating all formulas with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook and save it into a memory stream
                // ------------------------------------------------------------
                Workbook originalWorkbook = new Workbook();
                Worksheet sheet = originalWorkbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some initial data and a formula
                cells["A1"].PutValue(10);
                cells["B1"].Formula = "=A1*2"; // Original formula

                // Save the workbook to a MemoryStream (simulating an existing file)
                using (MemoryStream ms = new MemoryStream())
                {
                    originalWorkbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0; // Reset stream position for reading

                    // ------------------------------------------------------------
                    // 2. Load the workbook from the memory stream
                    // ------------------------------------------------------------
                    Workbook loadedWorkbook = new Workbook(ms);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                    Cells loadedCells = loadedSheet.Cells;

                    // ------------------------------------------------------------
                    // 3. Change the formula in cell B1
                    // ------------------------------------------------------------
                    loadedCells["B1"].Formula = "=A1*3"; // Updated formula

                    // ------------------------------------------------------------
                    // 4. Prepare calculation options (e.g., ignore errors)
                    // ------------------------------------------------------------
                    CalculationOptions calcOptions = new CalculationOptions
                    {
                        IgnoreError = true,   // Hide errors during calculation
                        Recursive = true      // Recalculate dependent cells across worksheets
                    };

                    // Optional: set iterative calculation at workbook level if needed
                    // Note: EnableIterativeCalculation property may not be available in all versions.
                    // loadedWorkbook.Settings.EnableIterativeCalculation = false;

                    // ------------------------------------------------------------
                    // 5. Calculate all formulas using the specified options
                    // ------------------------------------------------------------
                    try
                    {
                        loadedWorkbook.CalculateFormula(calcOptions);
                    }
                    catch (Exception calcEx)
                    {
                        Console.WriteLine("Calculation error: " + calcEx.Message);
                    }

                    // ------------------------------------------------------------
                    // 6. Output the result of the changed formula
                    // ------------------------------------------------------------
                    Console.WriteLine("Result of B1 after recalculation: " + loadedCells["B1"].Value);

                    // ------------------------------------------------------------
                    // 7. (Optional) Save the modified workbook back to a new memory stream
                    // ------------------------------------------------------------
                    using (MemoryStream outStream = new MemoryStream())
                    {
                        loadedWorkbook.Save(outStream, SaveFormat.Xlsx);
                        // The outStream now contains the updated workbook.
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
