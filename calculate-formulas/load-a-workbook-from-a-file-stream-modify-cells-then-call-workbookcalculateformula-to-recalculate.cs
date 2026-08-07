// Title: Load Excel Workbook from FileStream, Edit Cells, and Recalculate Formulas with Aspose.Cells for .NET
// Description: Shows how to open an XLSX file via a read‑only FileStream, create a Workbook, set a value in A1, assign a formula to B1, call Workbook.CalculateFormula to update all formulas, and read the computed result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells FileStream | load workbook from stream | Workbook.CalculateFormula | modify cell value C# | set formula C# | recalculate Excel formulas | read‑only stream Excel | Aspose.Cells .NET example
// Common Searches: Aspose.Cells load workbook from FileStream | C# calculate formulas after editing cells | Workbook.CalculateFormula example | how to update cell value and formula in Aspose.Cells | read‑only stream Excel Aspose.Cells tutorial
// Developer Intent: Open an Excel file from a stream, change cell data or formulas, and recalculate the workbook.
// Use Cases: Refresh financial calculations after importing a spreadsheet via API. | Automate data correction in uploaded Excel files before further processing. | Batch process large workbooks using streaming to reduce memory usage and recompute dependent formulas. | Generate reports where values are injected programmatically and formulas need immediate evaluation.
// AI Prompts: Write C# code that loads an XLSX file from a FileStream, sets A1 to a number, adds a formula to B1, runs Workbook.CalculateFormula, and returns B1's result using Aspose.Cells. | Show how to stream a large Excel workbook into Aspose.Cells, modify multiple cells, and efficiently recalculate all formulas without loading the whole file into memory. | Explain best practices for updating cells and formulas in a workbook opened from a read‑only stream, then persisting the calculated values.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to open an XLSX file via a read‑only FileStream, create a Workbook, set a value in A1, assign a formula to B1, call Workbook.CalculateFormula to update all formulas, and read the computed result using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFilePath = "input.xlsx";

        // Open a read‑only file stream
        using (FileStream stream = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
        {
            // Load the workbook from the stream (uses Workbook(Stream) constructor)
            Workbook workbook = new Workbook(stream);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify cells: set a value and a formula that depends on that value
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["B1"].Formula = "=A1*2";

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated result
            Console.WriteLine("Calculated value in B1: " + worksheet.Cells["B1"].IntValue);
        }
    }
}
