// Title: C# Aspose.Cells: Merge F4:G5, set a formula, calculate, and export to CSV
// Description: Load an XLSX workbook with Aspose.Cells for .NET, merge the range F4:G5, assign a value to the merged cell, create a formula in H4 that references it, evaluate all formulas, and save the worksheet as a CSV file.
// Keywords: Aspose.Cells merge cells C# | export CSV Aspose.Cells .NET | formula referencing merged cell | calculate formulas Aspose | Workbook.Save CSV example | C# spreadsheet manipulation
// Common Searches: Aspose.Cells merge range and export CSV | C# set formula after merging cells Aspose | how to calculate formulas before CSV save Aspose.Cells | merge cells F4:G5 Aspose.Cells .NET
// Developer Intent: Merge a specific cell block, apply a dependent formula, compute its result, and generate a CSV output using Aspose.Cells in C#.
// Use Cases: Prepare a CSV report where header cells are merged for visual clarity before export. | Create a calculated column that depends on a merged cell value for downstream data processing. | Automate spreadsheet-to-CSV conversion with pre‑evaluated formulas for analytics pipelines.
// AI Prompts: Write C# code with Aspose.Cells to merge cells F4:G5, add a formula that uses the merged cell, evaluate it, and save the sheet as CSV. | Explain the steps to calculate formulas after merging cells in Aspose.Cells and ensure the results appear in the exported CSV file. | Provide a concise tutorial for merging a range, inserting a formula, running calculations, and exporting to CSV using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an XLSX workbook with Aspose.Cells for .NET, merge the range F4:G5, assign a value to the merged cell, create a formula in H4 that references it, evaluate all formulas, and save the worksheet as a CSV file.
class MergeAndExportCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells F4:G5 (zero‑based indices: row 3, column 5, 2 rows, 2 columns)
        cells.Merge(3, 5, 2, 2);

        // Put a sample value into the merged cell for the formula to use
        cells["F4"].PutValue(10);

        // Insert a formula in H4 that references the merged cell (F4)
        cells["H4"].Formula = "=F4*2";

        // Calculate formulas so the result is stored
        workbook.CalculateFormula();

        // Export the worksheet to CSV
        string outputPath = "output.csv";
        workbook.Save(outputPath, SaveFormat.Csv);

        Console.WriteLine("Merge completed, formula set, and CSV exported successfully.");
    }
}
