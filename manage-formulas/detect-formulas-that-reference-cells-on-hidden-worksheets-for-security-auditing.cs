// Title: Detect Excel formulas that reference hidden worksheets with Aspose.Cells for .NET (C#)
// Common Searches: C# code to find Excel formulas that point to hidden sheets using Aspose.Cells | how to audit hidden worksheet references in Excel workbooks with Aspose.Cells .NET | scan all formulas for hidden sheet links in a workbook using Aspose.Cells C# | identify formulas referencing invisible worksheets in Excel via Aspose.Cells API | security check for hidden sheet references in Excel formulas with Aspose.Cells | list cells containing hidden sheet references in .xlsx using Aspose.Cells C#
// Tags: detect hidden sheet references Aspose.Cells | audit Excel formulas C# | scan workbook formulas for invisible worksheets | Aspose.Cells formula security analysis | enumerate hidden worksheets .NET | identify hidden worksheet links in Excel

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace HiddenSheetFormulaAudit
{
    // The example loads an Excel workbook with Aspose.Cells, gathers the names of all hidden worksheets, then iterates through every cell that contains a formula across all sheets. For each formula it checks whether the text includes a hidden sheet name followed by an exclamation mark, handling both quoted and unquoted sheet references, and logs any matches. This enables a security audit of formulas that reference hidden worksheets.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace the path with your actual file)
            Workbook workbook = new Workbook("input.xlsx");

            // Collect names of hidden worksheets
            HashSet<string> hiddenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.IsVisible == false) // Hidden worksheet
                {
                    hiddenSheetNames.Add(ws.Name);
                }
            }

            // Iterate through all cells with formulas and check if they reference hidden sheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Get the cells collection for the current worksheet
                Cells cells = ws.Cells;

                // Iterate over all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Simple detection: check if the formula string contains any hidden sheet name followed by '!'
                        foreach (string hiddenName in hiddenSheetNames)
                        {
                            // Formulas may reference sheets with or without single quotes
                            // e.g., Sheet1!A1 or 'My Hidden Sheet'!B2
                            string pattern1 = hiddenName + "!";
                            string pattern2 = $"'{hiddenName}'!";

                            if (cell.Formula.IndexOf(pattern1, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                cell.Formula.IndexOf(pattern2, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                Console.WriteLine($"Formula in sheet '{ws.Name}' cell {cell.Name} references hidden sheet '{hiddenName}': {cell.Formula}");
                                break; // No need to check other hidden sheets for this cell
                            }
                        }
                    }
                }
            }

            // Optionally, save a report workbook (replace the path as needed)
            // Here we just demonstrate saving the original workbook unchanged
            workbook.Save("output.xlsx");
        }
    }
}
