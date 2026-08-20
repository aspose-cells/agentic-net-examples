// Title: Remove a Worksheet and Export Remaining Sheets to PDF with Aspose.Cells (C#)
// Description: Shows how to delete a specific worksheet from an Aspose.Cells workbook using Worksheets.RemoveAt and then save the workbook as a PDF, guaranteeing the removed sheet is omitted from the final document.
// Keywords: Aspose.Cells | C# | remove worksheet | delete sheet | PDF export | exclude sheet from PDF | Worksheets.RemoveAt | Aspose.Cells .NET | remove sheet before PDF conversion
// Common Searches: Aspose.Cells delete worksheet before PDF export C# | How to exclude a sheet from PDF using Aspose.Cells | Remove specific worksheet and save as PDF in .NET | Worksheets.RemoveAt example for PDF generation
// Developer Intent: Delete a designated worksheet so it is not included in the generated PDF file.
// Use Cases: Create client‑specific PDFs that hide confidential worksheets. | Generate summary reports that contain only selected sheets. | Automate batch conversions where temporary sheets are stripped before PDF output.
// AI Prompts: Provide C# code that removes a worksheet by name with Aspose.Cells and then saves the workbook as a PDF. | Show how to delete multiple worksheets matching a pattern before exporting to PDF using Aspose.Cells. | Explain how to confirm that a removed worksheet does not appear in the resulting PDF.

using System;
using Aspose.Cells;

// Shows how to delete a specific worksheet from an Aspose.Cells workbook using Worksheets.RemoveAt and then save the workbook as a PDF, guaranteeing the removed sheet is omitted from the final document.
class RemoveWorksheetAndSavePdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Rename the default sheet and add two more sheets
        workbook.Worksheets[0].Name = "SheetToKeep";
        workbook.Worksheets.Add("SheetToRemove");
        workbook.Worksheets.Add("AnotherSheet");

        // Fill some data in each sheet
        workbook.Worksheets["SheetToKeep"].Cells["A1"].PutValue("This sheet will be kept");
        workbook.Worksheets["SheetToRemove"].Cells["A1"].PutValue("This sheet will be removed");
        workbook.Worksheets["AnotherSheet"].Cells["A1"].PutValue("This sheet will also be kept");

        // Remove the unwanted worksheet by its name
        workbook.Worksheets.RemoveAt("SheetToRemove");

        // Save the workbook as PDF; only the remaining sheets are rendered
        workbook.Save("Result.pdf", SaveFormat.Pdf);
    }
}
