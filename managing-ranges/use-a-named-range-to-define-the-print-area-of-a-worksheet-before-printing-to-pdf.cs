// Title: Set Print Area from a Named Range and Export to PDF with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a named range, assign its address to Worksheet.PageSetup.PrintArea, and save the workbook as a PDF so that only the defined range is printed using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# print area | named range PDF export | Worksheet.PageSetup.PrintArea | Aspose.Cells set print area | Excel to PDF Aspose.Cells | C# Aspose.Cells named range | PrintArea from named range | Aspose.Cells PDF conversion | Workbook.Save PDF | Aspose.Cells example
// Common Searches: Aspose.Cells set print area from named range C# | Export specific range to PDF using Aspose.Cells | How to use named range as print area in Aspose.Cells | C# Aspose.Cells print only selected cells to PDF | PageSetup.PrintArea named range example
// Developer Intent: Define a worksheet's print area using a named range and generate a PDF that contains only that area.
// Use Cases: Create a reusable named range to control which cells are printed when exporting to PDF. | Dynamically change the RefersTo address of a named range to adjust the printed region without modifying code. | Apply the same named‑range‑based print area across multiple worksheets for batch PDF generation.
// AI Prompts: Write C# code with Aspose.Cells that defines a named range, sets Worksheet.PageSetup.PrintArea to that range, and saves the workbook as a PDF. | Explain how to update an existing named range's RefersTo property and refresh the PrintArea for PDF export in Aspose.Cells. | Provide step‑by‑step instructions for using a named range as a print area, including handling the leading '=' in the RefersTo string.

using System;
using Aspose.Cells;

// Demonstrates how to create a named range, assign its address to Worksheet.PageSetup.PrintArea, and save the workbook as a PDF so that only the defined range is printed using Aspose.Cells for .NET.
class PrintAreaFromNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item2");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("Item3");
        worksheet.Cells["B4"].PutValue(300);

        // Create a named range that will be used as the print area (A1:B3)
        int nameIdx = workbook.Worksheets.Names.Add("PrintRange");
        Name namedRange = workbook.Worksheets.Names[nameIdx];
        // RefersTo must start with '=' and include the sheet name
        namedRange.RefersTo = $"={worksheet.Name}!A1:B3";

        // Set the worksheet's print area to the address of the named range
        // Remove the leading '=' from RefersTo as PrintArea expects a plain address
        worksheet.PageSetup.PrintArea = namedRange.RefersTo.Substring(1);

        // Save the workbook to PDF; only the defined print area will be exported
        workbook.Save("PrintAreaFromNamedRange.pdf");
    }
}
