// Title: C# – Remove All External Links from an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file, clears every external link (optionally updating formulas to refer to the current workbook), and saves a self‑contained workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# external links | clear external links | remove external references | self-contained Excel file | update formulas after clearing links | Workbook.ExternalLinks.Clear | Excel data connections | .NET Excel security
// Common Searches: Aspose.Cells remove external links C# | How to clear external references in Excel using .NET | Make Excel workbook self‑contained with Aspose | Workbook.ExternalLinks.Clear example | Delete external data connections programmatically
// Developer Intent: Delete every external link in a workbook so the file no longer depends on outside sources.
// Use Cases: Prepare a distribution‑ready workbook that must not contain external data connections. | Enforce compliance by stripping external references from corporate templates. | Automate conversion of user‑uploaded Excel files into secure, self‑contained documents.
// AI Prompts: Write C# code with Aspose.Cells that lists all external links before removing them. | Show how to call Workbook.Worksheets.ExternalLinks.Clear(false) to keep original formulas unchanged. | Explain how to verify that no external links remain after calling Clear(true).

using System;
using Aspose.Cells;

// Loads an Excel file, clears every external link (optionally updating formulas to refer to the current workbook), and saves a self‑contained workbook using Aspose.Cells for .NET.
class RemoveExternalLinksDemo
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all external links.
        // The boolean parameter updates formulas to refer to the current workbook when possible.
        workbook.Worksheets.ExternalLinks.Clear(true);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
