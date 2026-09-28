// Title: Repeat a header row on each PDF page when exporting an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that sets the first worksheet row as a print title so the header repeats on every PDF page, then saves the workbook as a PDF. | Show how to define a custom print area and configure page‑setup options (print titles, margins) before converting an Excel file to PDF with Aspose.Cells in C#. | Provide an example that creates sample data, applies a repeat‑header setting, and exports the sheet to a PDF using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export Excel to PDF with repeating header row on each page | set print title rows in Aspose.Cells before PDF conversion .NET | how to define print area and repeat header when saving workbook as PDF using Aspose.Cells | C# Aspose.Cells page setup repeat header on PDF pages | export worksheet to PDF with print titles using Aspose.Cells for .NET
// Tags: Aspose.Cells set print title rows C# | Aspose.Cells define print area PDF export | repeat header rows PDF generation Aspose.Cells | page setup configuration Aspose.Cells .NET | export Excel worksheet to PDF Aspose.Cells

using Aspose.Cells;
using System;

// // Creates a workbook, adds a header and sample rows, sets the first row as a print title to repeat on every PDF page, defines the print area, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        // Populate sample data rows
        for (int i = 2; i <= 20; i++)
        {
            sheet.Cells[i - 1, 0].PutValue(i - 1);               // ID
            sheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");    // Name
            sheet.Cells[i - 1, 2].PutValue(50 + i);             // Score
        }

        // Set the first row as print title (repeat on each PDF page)
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // Define the print area to include all data
        sheet.PageSetup.PrintArea = "$A$1:$C$20";

        // Save the workbook as PDF
        workbook.Save("Output.pdf", SaveFormat.Pdf);
    }
}
