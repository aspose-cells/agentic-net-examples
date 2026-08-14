// Title: Aspose.Cells .NET Example: Generate XLSX File with a Centered Header (C#)
// Description: This C# snippet shows how to create a Workbook, configure PageSetup to place custom text in the center of the header, and save the result as an XLSX document using Aspose.Cells.
// Keywords: Aspose.Cells C# header | centered header Excel | PageSetup SetHeader | save workbook XLSX | Excel printing header | C# Aspose.Cells example | create workbook with header
// Common Searches: Aspose.Cells set center header C# | How to add a header to Excel file using Aspose.Cells .NET | PageSetup SetHeader example | Save workbook as XLSX with header Aspose | C# code for Excel header formatting
// Developer Intent: Add custom text to the middle section of a worksheet header and export the sheet as an XLSX file.
// Use Cases: Print‑ready reports where the title appears centered on each page | Standardized templates that embed a company name in the header | Automated invoices or statements with a centered header containing document ID | Batch generation of spreadsheets that require a consistent header layout
// AI Prompts: Generate C# code that creates an XLSX workbook, sets a centered header using Aspose.Cells, and saves the file. | Provide a step‑by‑step guide to apply the same centered header to all worksheets in a workbook with Aspose.Cells. | Explain how to combine left, center, and right header sections in Aspose.Cells PageSetup for a multi‑page report.

using System;
using Aspose.Cells;

namespace CenteredHeaderExample
{
    // This C# snippet shows how to create a Workbook, configure PageSetup to place custom text in the center of the header, and save the result as an XLSX document using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the centered header (section 1) with the desired text
            pageSetup.SetHeader(1, "My Centered Header");

            // Save the workbook to an XLSX file
            workbook.Save("CenteredHeader.xlsx");

            Console.WriteLine("Workbook created with centered header and saved as CenteredHeader.xlsx");
        }
    }
}
