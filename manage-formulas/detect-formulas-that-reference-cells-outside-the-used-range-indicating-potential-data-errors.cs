// Title: Detect Out‑of‑Range Formula References in Excel with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds data, inserts valid and invalid formulas, calculates the sheet, obtains the used range via MaxDataRow/MaxDataColumn, scans all formula cells with GetPrecedents, and reports any reference that lies outside the defined data area.
// Keywords: Aspose.Cells | .NET | C# | detect out of range formulas | GetPrecedents | used range | MaxDataRow | MaxDataColumn | Excel data validation | formula reference audit
// Common Searches: Aspose.Cells find formulas outside used range | C# detect invalid cell references in Excel | GetPrecedents out of range check Aspose.Cells | How to validate formula references with Aspose.Cells .NET | Identify #REF! errors programmatically
// Developer Intent: Locate and list formula cells that reference rows or columns beyond the worksheet’s current used range.
// Use Cases: Validate workbook integrity by flagging formulas that point to empty or non‑existent cells. | Generate an audit report of out‑of‑range references before publishing or sharing the file. | Automate cleanup workflows that correct or remove formulas exceeding defined data boundaries.
// AI Prompts: Write a C# method using Aspose.Cells that returns all cells whose formulas reference rows or columns beyond MaxDataRow or MaxDataColumn. | Provide code to replace out‑of‑range references in formulas with #REF! using Aspose.Cells. | Explain how GetPrecedents and ReferredArea can be used to separate external links from internal out‑of‑range references.

using System;
using Aspose.Cells;

// Creates a workbook, adds data, inserts valid and invalid formulas, calculates the sheet, obtains the used range via MaxDataRow/MaxDataColumn, scans all formula cells with GetPrecedents, and reports any reference that lies outside the defined data area.
class DetectOutOfRangeFormulas
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some data within a normal used range (A1:B2)
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].PutValue(30);
        cells["B2"].PutValue(40);

        // Valid formula that stays inside the used range
        cells["C1"].Formula = "=SUM(A1:B2)";

        // Invalid formula that references cells outside the used range
        cells["D1"].Formula = "=E5+F6";

        // Optional: calculate formulas so that dependent values are up‑to‑date
        workbook.CalculateFormula();

        // Determine the current used range of the worksheet
        int maxDataRow = cells.MaxDataRow;       // zero‑based index of the last row with data
        int maxDataColumn = cells.MaxDataColumn; // zero‑based index of the last column with data

        // Scan all cells that contain formulas
        foreach (Cell cell in cells)
        {
            if (string.IsNullOrEmpty(cell.Formula))
                continue; // Skip non‑formula cells

            // Get all precedent areas referenced by the formula
            ReferredAreaCollection precedents = cell.GetPrecedents();
            if (precedents == null)
                continue; // No precedents (should not happen for a formula)

            foreach (ReferredArea area in precedents)
            {
                // Ignore external links; we only care about internal references
                if (area.IsExternalLink)
                    continue;

                // Determine if any part of the referenced area lies outside the used range
                bool outOfRow = area.StartRow > maxDataRow || area.EndRow > maxDataRow;
                bool outOfColumn = area.StartColumn > maxDataColumn || area.EndColumn > maxDataColumn;

                if (outOfRow || outOfColumn)
                {
                    Console.WriteLine($"Formula in cell {cell.Name} references out‑of‑range area: {area}");
                }
            }
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("DetectedOutOfRangeFormulas.xlsx");
    }
}
