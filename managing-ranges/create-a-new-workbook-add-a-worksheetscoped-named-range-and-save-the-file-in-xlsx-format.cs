// Title: C# – Create a Worksheet‑Scoped Named Range in a New XLSX Workbook with Aspose.Cells
// Description: Demonstrates how to instantiate a Workbook, rename the first worksheet, define a range covering A1:B3, assign a worksheet‑scoped name using the "SheetName!RangeName" syntax, and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# example | worksheet scoped named range | create range A1:B3 | save workbook as XLSX | SheetName!RangeName syntax | Excel automation .NET | Aspose.Cells named range tutorial | C# Excel library
// Common Searches: how to add a worksheet scoped named range with Aspose.Cells | Aspose.Cells create named range and save as xlsx | C# code for worksheet‑specific named range in Excel | Aspose.Cells range naming conventions | save workbook with named range using Aspose.Cells .NET
// Developer Intent: Generate an XLSX workbook, define a worksheet‑specific named range, and persist the file.
// Use Cases: Reference a data block on a single sheet without affecting similarly named ranges on other sheets. | Automate report templates where each worksheet contains its own scoped range for downstream processing. | Provide end‑users with pre‑named areas for data entry, ensuring consistent formula references across multiple sheets.
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a worksheet‑scoped named range covering C5:D10, and saves it as XLSX. | Explain the "SheetName!RangeName" format in Aspose.Cells and show how to retrieve a worksheet‑scoped range later in C#. | Give a step‑by‑step guide to create several worksheet‑scoped named ranges in one workbook and use them in formulas with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeDemo
{
    // Demonstrates how to instantiate a Workbook, rename the first worksheet, define a range covering A1:B3, assign a worksheet‑scoped name using the "SheetName!RangeName" syntax, and save the file as an XLSX document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "DataSheet";

                // Create a range on the worksheet (cells A1 to B3)
                AsposeRange range = worksheet.Cells.CreateRange("A1:B3");

                // Assign a worksheet‑scoped name to the range.
                // The "SheetName!RangeName" format makes the name scoped to this worksheet.
                range.Name = $"{worksheet.Name}!MyWorksheetRange";

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
