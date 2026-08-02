// Title: C# – Create a Worksheet‑Scoped Named Range and Save as XLSX with Aspose.Cells
// Description: Demonstrates how to instantiate a new Workbook, rename the first worksheet, populate cells A1:B3, define a range A2:B3, assign a worksheet‑scoped name (e.g., DataSheet!MyRange), and save the file as WorksheetScopedNamedRange.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet scoped named range | C# create named range Aspose.Cells | save workbook as xlsx Aspose.Cells | Aspose.Cells range name prefix sheet | Aspose.Cells .NET example
// Common Searches: Aspose.Cells create worksheet scoped named range C# | How to save a workbook with a named range in XLSX using Aspose.Cells | Define a range on a specific sheet and give it a name in .NET | Worksheet‑scoped named range Aspose.Cells tutorial
// Developer Intent: Generate a new XLSX workbook, add a worksheet‑scoped named range, and persist the file.
// Use Cases: Reference a fixed data block in formulas or external tools without ambiguity. | Provide a ready‑made template where downstream processes locate a named range on a particular sheet. | Enable data validation, chart sources, or macro‑driven automation that rely on sheet‑specific named ranges.
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, fills cells A1:B3, defines a worksheet‑scoped named range covering A2:B3, and saves it as XLSX. | Explain why a worksheet‑scoped named range requires the sheet name prefix and how Aspose.Cells implements it. | Give step‑by‑step instructions for adding multiple worksheet‑scoped named ranges across different sheets in the same workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetScopedNamedRange
{
    // Demonstrates how to instantiate a new Workbook, rename the first worksheet, populate cells A1:B3, define a range A2:B3, assign a worksheet‑scoped name (e.g., DataSheet!MyRange), and save the file as WorksheetScopedNamedRange.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a meaningful name
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "DataSheet";

                // Populate some sample data (optional)
                worksheet.Cells["A1"].PutValue("Item");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(15);

                // Create a range that will be named (use fully qualified Aspose.Cells.Range to avoid ambiguity)
                Aspose.Cells.Range namedRange = worksheet.Cells.CreateRange("A2:B3");

                // Assign a worksheet‑scoped name (prefix with the sheet name)
                namedRange.Name = $"{worksheet.Name}!MyRange";

                // Save the workbook in XLSX format
                workbook.Save("WorksheetScopedNamedRange.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
