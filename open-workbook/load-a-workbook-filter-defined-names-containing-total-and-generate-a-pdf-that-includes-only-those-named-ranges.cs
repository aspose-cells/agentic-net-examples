// Title: Export Only Sheets with "Total" Named Ranges to PDF using Aspose.Cells for .NET
// Description: C# code that loads an Excel workbook, filters all defined names (workbook‑ and worksheet‑scoped) for the keyword "Total", builds a SheetSet of the matching sheet indexes, and saves those sheets as a single PDF with Aspose.Cells.
// Keywords: Aspose.Cells PDF export | C# defined names filter | named range PDF Aspose | SheetSet save options | export selected worksheets | filter named ranges Total | Aspose.Cells .NET tutorial | Excel to PDF selective sheets | workbook scoped names | worksheet scoped names
// Common Searches: Aspose.Cells export only sheets with specific named range | C# filter defined names containing Total and save as PDF | How to use SheetSet in Aspose.Cells to create a PDF | Select worksheets by named range keyword Aspose.Cells | Save part of a workbook to PDF with Aspose.Cells .NET
// Developer Intent: Generate a PDF that includes only the worksheets whose defined names contain the word "Total".
// Use Cases: Create a financial summary PDF that contains only sheets with total calculations. | Automate reporting by exporting only relevant worksheets to a compact PDF. | Produce an audit‑ready document that excludes sheets without "Total" named ranges.
// AI Prompts: Modify the example to also include named ranges that start with "Total" but ignore those that only contain the word. | Show how to write the filtered PDF to a MemoryStream instead of a file path. | Explain handling of workbook‑scoped named ranges that map to the first worksheet when building the SheetSet.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsDefinedNamesPdf
{
    // C# code that loads an Excel workbook, filters all defined names (workbook‑ and worksheet‑scoped) for the keyword "Total", builds a SheetSet of the matching sheet indexes, and saves those sheets as a single PDF with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Get all defined names (both workbook‑scoped and worksheet‑scoped)
            Name[] allNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

            // Collect distinct sheet indexes that contain names with "Total"
            HashSet<int> sheetIndexes = new HashSet<int>();
            foreach (Name name in allNames)
            {
                if (name.Text != null && name.Text.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // SheetIndex = 0 for global names, otherwise 1‑based sheet index
                    // Convert to zero‑based index for SheetSet
                    int zeroBasedIndex = name.SheetIndex == 0 ? 0 : name.SheetIndex - 1;
                    sheetIndexes.Add(zeroBasedIndex);
                }
            }

            // If no matching names were found, exit
            if (sheetIndexes.Count == 0)
            {
                Console.WriteLine("No defined names containing 'Total' were found.");
                return;
            }

            // Prepare PDF save options and limit the output to the selected sheets
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Create a SheetSet from the collected sheet indexes
            SheetSet sheetSet = new SheetSet(new List<int>(sheetIndexes).ToArray());
            pdfOptions.SheetSet = sheetSet;

            // Save the workbook as PDF containing only the relevant sheets
            workbook.Save("FilteredTotalRanges.pdf", pdfOptions);

            Console.WriteLine("PDF generated successfully with sheets that contain 'Total' named ranges.");
        }
    }
}
