// Title: C# – Define a Named Range for a Dynamic Array Spill Range using Aspose.Cells
// Description: This example creates a workbook, writes source values, inserts a SEQUENCE dynamic‑array formula, calculates and refreshes it, retrieves the spill area with GetArrayRange, builds a Range object for that area, and registers a workbook‑level named range that points to the spill range before saving the file.
// Keywords: Aspose.Cells dynamic array | spill range GetArrayRange | C# named range from dynamic array | SEQUENCE formula Aspose.Cells | CreateRange from CellArea | RefreshDynamicArrayFormulas | Aspose.Cells .NET example
// Common Searches: Aspose.Cells create named range for dynamic array spill | GetArrayRange example C# | How to reference SEQUENCE spill range in Aspose.Cells | Define workbook name for dynamic array output | C# Aspose.Cells refresh dynamic array formulas
// Developer Intent: Create a workbook‑level named range that automatically tracks the spill area produced by a dynamic‑array formula.
// Use Cases: Link charts to a named range that expands as the SEQUENCE output grows. | Use the named range in other formulas or external reports to ensure consistent data references. | Export workbooks to systems that rely on named ranges for locating variable‑size data blocks.
// AI Prompts: Provide C# code that defines a named range for the spill area of a SEQUENCE dynamic array using Aspose.Cells. | Show how to refresh dynamic‑array formulas, obtain the spill range with GetArrayRange, and assign it to a workbook name. | Explain the steps to convert a CellArea to a Range object and set the RefersTo property for a named range in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicArrayNamedRange
{
    // This example creates a workbook, writes source values, inserts a SEQUENCE dynamic‑array formula, calculates and refreshes it, retrieves the spill area with GetArrayRange, builds a Range object for that area, and registers a workbook‑level named range that points to the spill range before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some data that the dynamic array formula will use
                cells["B1"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["B3"].PutValue(30);

                // Set a dynamic array formula in cell A1 (e.g., SEQUENCE based on B1 value)
                Cell startCell = cells["A1"];
                startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

                // Calculate formulas and refresh the spill range
                workbook.CalculateFormula();
                workbook.RefreshDynamicArrayFormulas(true);

                // Get the actual spill range of the dynamic array formula
                CellArea spillArea = startCell.GetArrayRange();

                // Create a Range object that represents the spill area
                int rows = spillArea.EndRow - spillArea.StartRow + 1;
                int cols = spillArea.EndColumn - spillArea.StartColumn + 1;
                AsposeRange spillRange = cells.CreateRange(spillArea.StartRow, spillArea.StartColumn, rows, cols);

                // Define a named range that refers to the spill range
                int nameIdx = workbook.Worksheets.Names.Add("MyDynamicArray");
                // RefersTo must start with '=' and include the sheet name
                workbook.Worksheets.Names[nameIdx].RefersTo = "=" + spillRange.RefersTo;

                // Optional: verify the named range address
                Console.WriteLine("Named range 'MyDynamicArray' refers to: " + workbook.Worksheets.Names[nameIdx].RefersTo);

                // Save the workbook
                string outputPath = "DynamicArrayNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
