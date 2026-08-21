// Title: Copy a Range with Data, Formulas & Styles to Another Worksheet using PasteOptions – Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, fills a source sheet with headers, values, formulas and a bold gray style, then uses Aspose.Cells PasteOptions (PasteType.All) to duplicate the A1:D3 range—including values, formulas, and formatting—into a new worksheet and saves the file.
// Keywords: Aspose.Cells copy range C# | PasteOptions duplicate range | preserve formulas Aspose.Cells | copy cell styles .NET | Excel range copy example | Aspose.Cells tutorial | GitHub Aspose.Cells sample | C# Excel automation
// Common Searches: Aspose.Cells copy range with formulas C# | How to duplicate a range with formatting using PasteOptions | Copy Excel cells to another sheet preserving styles Aspose.Cells | C# PasteOptions PasteType.All example | Aspose.Cells range copy to new worksheet
// Developer Intent: Duplicate a source range—including values, formulas, and formatting—to a destination range on a new worksheet using PasteOptions.
// Use Cases: Generate multiple report tabs by cloning a styled, calculated table from a template sheet. | Create a summary sheet that reuses a formatted data block while keeping all underlying formulas intact. | Export a fully formatted and calculated table to a separate workbook for distribution without altering the original.
// AI Prompts: Provide C# code that copies a range with formulas and cell styles to another worksheet using Aspose.Cells PasteOptions. | Show how to use PasteOptions to copy a range while skipping blank cells but retaining formatting in Aspose.Cells for .NET. | Explain how to transpose a range during copy while preserving formulas and styles with Aspose.Cells PasteOptions. | Give a step‑by‑step guide to duplicate a range to a new worksheet and save the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsRangeDuplicate
{
    // C# example that creates a workbook, fills a source sheet with headers, values, formulas and a bold gray style, then uses Aspose.Cells PasteOptions (PasteType.All) to duplicate the A1:D3 range—including values, formulas, and formatting—into a new worksheet and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------- Source Worksheet --------------------
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Populate source range with data, formulas and style
                srcSheet.Cells["A1"].PutValue("Item");
                srcSheet.Cells["B1"].PutValue("Quantity");
                srcSheet.Cells["C1"].PutValue("Price");
                srcSheet.Cells["A2"].PutValue("Apple");
                srcSheet.Cells["B2"].PutValue(10);
                srcSheet.Cells["C2"].PutValue(0.5);
                srcSheet.Cells["A3"].PutValue("Banana");
                srcSheet.Cells["B3"].PutValue(5);
                srcSheet.Cells["C3"].PutValue(0.3);

                // Formula (Total = Quantity * Price)
                srcSheet.Cells["D2"].Formula = "B2*C2";
                srcSheet.Cells["D3"].Formula = "B3*C3";

                // Apply a simple style to the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;
                srcSheet.Cells.CreateRange("A1:D1").SetStyle(headerStyle);

                // -------------------- Destination Worksheet --------------------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Define source and destination ranges (same size)
                Aspose.Cells.Range sourceRange = srcSheet.Cells.CreateRange("A1:D3");
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:D3");

                // Configure paste options to copy everything (data, formulas, formats, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All,   // copy all aspects
                    SkipBlanks = false,         // do not skip blanks
                    Transpose = false           // keep original orientation
                };

                // Perform the copy using the range copy method with paste options
                destRange.Copy(sourceRange, pasteOptions);

                // Save the workbook
                workbook.Save("RangeDuplicateWithCopyOptions.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
