// Title: Merge Cells F4:G5, Apply a Formula, and Export to CSV using Aspose.Cells for .NET
// Description: Load an Excel workbook, merge the range F4:G5 on the first worksheet, put a value, set a formula that references the merged cell, recalculate all formulas, and save the result as a CSV file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | Aspose.Cells CSV export | Aspose.Cells formula reference | C# merge cells | C# export to CSV | Aspose.Cells calculate formulas | Aspose.Cells .NET | Excel to CSV conversion | merged cell reference | worksheet.Save CSV
// Common Searches: Aspose.Cells merge range F4:G5 C# | How to reference a merged cell in a formula with Aspose.Cells | Export Excel worksheet with merged cells to CSV using Aspose.Cells | Recalculate formulas after merging cells Aspose.Cells .NET | Save workbook as CSV after applying formulas Aspose.Cells
// Developer Intent: Merge a specific cell range, use it in a formula, recalculate, and generate a CSV file.
// Use Cases: Create a CSV report where a calculated value depends on a merged header cell. | Automate data preparation by merging cells, applying business logic, and exporting the result for downstream processing. | Convert an Excel template that contains merged cells and formulas into a flat CSV format for import into other systems.
// AI Prompts: Generate C# code with Aspose.Cells that merges cells F4:G5, inserts a value, adds a formula referencing the merged cell, recalculates, and saves the worksheet as CSV. | Explain how Aspose.Cells handles formula evaluation for merged cells when exporting to CSV format. | Provide troubleshooting steps if the calculated result of a formula referencing a merged cell does not appear in the exported CSV.

using System;
using Aspose.Cells;

// Load an Excel workbook, merge the range F4:G5 on the first worksheet, put a value, set a formula that references the merged cell, recalculate all formulas, and save the result as a CSV file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source workbook (modify as needed)
        string inputFile = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells F4:G5.
        // F4 -> row 3, column 5 (zero‑based)
        // The range spans 2 rows and 2 columns.
        worksheet.Cells.Merge(3, 5, 2, 2);

        // Optional: put a sample value into the merged cell.
        worksheet.Cells["F4"].PutValue(10);

        // Insert a formula that references the merged cell.
        // Example: place the formula in H6 (row 5, column 7).
        worksheet.Cells["H6"].Formula = "=F4*2";

        // Recalculate formulas so the result is stored.
        workbook.CalculateFormula();

        // Export the worksheet to CSV.
        string outputFile = "output.csv";
        workbook.Save(outputFile, SaveFormat.Csv);
    }
}
