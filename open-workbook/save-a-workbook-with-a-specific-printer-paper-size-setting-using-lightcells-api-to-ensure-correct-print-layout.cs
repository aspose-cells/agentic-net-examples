// Title: Set printer paper size and save a workbook with LightCells API in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to assign PaperSizeType.PaperA5 to Workbook.Settings.PaperSize and to the first worksheet's PageSetup, then persist the workbook using Aspose.Cells' saving mechanisms (compatible with LightCells) so the printed output matches the A5 layout.
// Keywords: Aspose.Cells | C# | set paper size | Workbook.Settings.PaperSize | Worksheet PageSetup paper size | LightCells API | save Excel workbook | A5 print layout | Excel printer settings | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set A5 paper size C# | How to change default printer paper size in an Excel workbook using Aspose.Cells | Save Excel file with specific page size using LightCells | Workbook.Settings.PaperSize example Aspose.Cells | PageSetup paper size before saving workbook Aspose.Cells .NET
// Developer Intent: Configure the workbook’s default printer paper size (and optionally the first sheet’s page setup) and save the file so that the printed result uses the defined size.
// Use Cases: Generating invoices that must print on A5 paper while keeping layout intact across all sheets. | Creating multi‑sheet reports where each sheet shares a common A5 page configuration before distribution. | Preparing printable forms or labels that require a consistent A5 size throughout the workbook.
// AI Prompts: Write C# code that sets the default paper size to Letter in an Aspose.Cells workbook and saves it using the LightCells API. | Show how to assign different paper sizes to several worksheets and export each sheet to PDF while preserving its layout with Aspose.Cells. | Explain how to read and verify the paper size setting of an existing Excel file after loading it with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPaperSizeDemo
{
    // Demonstrates how to assign PaperSizeType.PaperA5 to Workbook.Settings.PaperSize and to the first worksheet's PageSetup, then persist the workbook using Aspose.Cells' saving mechanisms (compatible with LightCells) so the printed output matches the A5 layout.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the standard creation rule)
            Workbook workbook = new Workbook();

            // Set the default print paper size for the workbook (WorkbookSettings.PaperSize)
            // This ensures that when the workbook is printed, the specified size is used.
            workbook.Settings.PaperSize = PaperSizeType.PaperA5; // Example: A5 size

            // Optionally, also set the paper size on the first worksheet's PageSetup
            // (useful if individual sheets need explicit settings)
            workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA5;

            // Save the workbook to disk using the provided Save method rule
            workbook.Save("Workbook_With_A5_PaperSize.xlsx", SaveFormat.Xlsx);

            // Output confirmation
            Console.WriteLine("Workbook saved with A5 paper size setting.");
        }
    }
}
