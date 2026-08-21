// Title: Audit Excel Formulas for Hidden Worksheet References with Aspose.Cells (C#)
// Description: Load an Excel workbook, identify all hidden worksheets, scan every formula cell using GetPrecedents, and report any formulas that reference hidden sheets. Ideal for security and compliance audits, with optional saving of the audited file.
// Keywords: Aspose.Cells | C# | hidden worksheet detection | Excel formula audit | GetPrecedents | security audit Excel | hidden sheet reference | workbook compliance | cell precedent analysis | .NET Excel processing
// Common Searches: Aspose.Cells detect formulas referencing hidden sheets | C# audit Excel workbook for hidden worksheet references | GetPrecedents hidden sheet detection example | How to find hidden sheet dependencies in Excel using Aspose.Cells | Security audit Excel formulas hidden worksheets C#
// Developer Intent: Find and list all formula cells that depend on hidden worksheets in an Excel file using Aspose.Cells for .NET.
// Use Cases: Perform a security review to ensure no formulas expose data from hidden sheets before sharing a workbook. | Validate regulatory compliance by confirming that published Excel files contain no hidden‑sheet calculations. | Generate a detailed report of cell addresses that reference hidden worksheets for debugging or documentation.
// AI Prompts: Create a C# method with Aspose.Cells that returns a list of cell addresses whose formulas reference hidden worksheets. | Show how to export the hidden‑sheet reference report to CSV using Aspose.Cells in a .NET application. | Explain step‑by‑step how GetPrecedents can be leveraged to detect hidden worksheet dependencies in Excel formulas.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Load an Excel workbook, identify all hidden worksheets, scan every formula cell using GetPrecedents, and report any formulas that reference hidden sheets. Ideal for security and compliance audits, with optional saving of the audited file.
class DetectHiddenSheetReferences
{
    static void Main()
    {
        // Load the workbook to be audited
        Workbook workbook = new Workbook("input.xlsx");

        // Collect names of all hidden worksheets
        HashSet<string> hiddenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // In Aspose.Cells, a worksheet is hidden when IsVisible is false
            if (!ws.IsVisible)
            {
                hiddenSheetNames.Add(ws.Name);
            }
        }

        // Scan every cell in every worksheet for formulas that reference hidden sheets
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Cells cells = ws.Cells;

            // Enumerate all cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Process only formula cells
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Get all precedent areas referenced by this formula
                    ReferredAreaCollection precedents = cell.GetPrecedents();

                    if (precedents != null)
                    {
                        foreach (ReferredArea area in precedents)
                        {
                            // The sheet name that the area refers to
                            string referencedSheet = area.SheetName;

                            // If the referenced sheet is hidden, report it
                            if (hiddenSheetNames.Contains(referencedSheet))
                            {
                                Console.WriteLine(
                                    $"Formula in {ws.Name}!{cell.Name} references hidden sheet '{referencedSheet}'.");
                            }
                        }
                    }
                }
            }
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("audit_output.xlsx");
    }
}
