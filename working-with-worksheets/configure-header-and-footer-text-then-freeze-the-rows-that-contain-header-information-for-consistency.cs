// Title: Configure left, center, right headers and footers, freeze the top row, and repeat header on printed pages with Aspose.Cells (C#)
// Description: This example creates a workbook, assigns custom text to the left, center, and right sections of the header and footer via PageSetup, freezes the first worksheet row using FreezePanes, sets the header row to repeat on every printed page with PrintTitleRows, and saves the file as HeaderFooterFreezeDemo.xlsx.
// Keywords: Aspose.Cells header C# | Aspose.Cells footer C# | freeze first row Aspose.Cells | repeat header on print Aspose.Cells | PageSetup SetHeader SetFooter | FreezePanes C# Aspose.Cells | PrintTitleRows Aspose.Cells
// Common Searches: Aspose.Cells set left header C# | How to freeze top row in Aspose.Cells | Repeat header row on each printed page Aspose.Cells | Add page numbers to footer Aspose.Cells C# | Configure three‑section header Aspose.Cells
// Developer Intent: Add custom text to the three header sections and footer, keep the header row visible while scrolling, and ensure it repeats on every printed page using Aspose.Cells for .NET.
// Use Cases: Insert company name, report title, and date into left, center, and right header sections of an automated Excel report. | Freeze the header row so it remains in view when users scroll through large data sets. | Automatically repeat the header row on each printed page for multi‑page spreadsheets.
// AI Prompts: Generate C# code with Aspose.Cells that sets left, center, and right header text, adds a footer with page numbers, freezes the first row, and repeats that row on printed pages. | Show how to use PageSetup.SetHeader, SetFooter, FreezePanes, and PrintTitleRows together in an Aspose.Cells .NET example. | Provide a complete Aspose.Cells C# snippet that configures three‑section headers, a numbered footer, freezes the top row, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderFooterFreezeDemo
{
    // This example creates a workbook, assigns custom text to the left, center, and right sections of the header and footer via PageSetup, freezes the first worksheet row using FreezePanes, sets the header row to repeat on every printed page with PrintTitleRows, and saves the file as HeaderFooterFreezeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate some sample data (optional, for demo)
            // -------------------------------------------------
            worksheet.Cells["A1"].PutValue("Header");
            for (int i = 2; i <= 20; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data {i - 1}");
            }

            // -------------------------------------------------
            // Configure header text (left, center, right sections)
            // -------------------------------------------------
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.SetHeader(0, "Left Header");               // Section 0 = Left
            pageSetup.SetHeader(1, "Center Header");             // Section 1 = Center
            pageSetup.SetHeader(2, "Right Header");              // Section 2 = Right

            // -------------------------------------------------
            // Configure footer text (left, center, right sections)
            // -------------------------------------------------
            pageSetup.SetFooter(0, "Left Footer");
            pageSetup.SetFooter(1, "Page &P of &N"); // Center shows page number
            pageSetup.SetFooter(2, "Right Footer");

            // -------------------------------------------------
            // Freeze the first row so header stays visible while scrolling
            // -------------------------------------------------
            // Freeze panes at cell A2, freezing 1 row and 0 columns
            worksheet.FreezePanes("A2", 1, 0);

            // -------------------------------------------------
            // Ensure the header row repeats on each printed page
            // -------------------------------------------------
            pageSetup.PrintTitleRows = "$1:$1";

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("HeaderFooterFreezeDemo.xlsx");
        }
    }
}
