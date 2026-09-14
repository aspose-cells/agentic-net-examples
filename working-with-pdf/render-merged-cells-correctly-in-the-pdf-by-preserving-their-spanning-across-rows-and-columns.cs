// Title: How to export a worksheet with merged rows and columns to PDF while preserving the merge using Aspose.Cells for C#
// AI Prompts: Write C# code that merges a header across multiple columns and a label across multiple rows, then saves the worksheet to PDF using Aspose.Cells with options that keep the merged cells intact. | Demonstrate how to configure PdfSaveOptions in Aspose.Cells to prevent page breaks from splitting merged cells and to force all columns onto a single PDF page. | Show how to apply center alignment to merged cells in a workbook before exporting to PDF with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# export merged cells to PDF without losing span | keep merged cell layout when saving Excel as PDF using Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet effect on merged cells Aspose | C# merge cells A1:C1 and A2:A4 then convert to PDF with Aspose.Cells | prevent page break inside merged cells Aspose.Cells PDF export
// Tags: Aspose.Cells merge cells export PDF | PdfSaveOptions preserve merged layout | C# center alignment merged cells Aspose | AllColumnsInOnePagePerSheet merged cells | prevent page break merged cells Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, merges cells A1:C1 and A2:A4, centers the text in those merged cells, configures PdfSaveOptions with OnePagePerSheet = false and AllColumnsInOnePagePerSheet = true, and saves the worksheet as a PDF, ensuring the merged cells retain their spanning in the output.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Row1");
        sheet.Cells["B2"].PutValue("Data1");
        sheet.Cells["C2"].PutValue("Data2");

        // Merge cells A1:C1 so the header spans three columns
        // Parameters: start row, start column, total rows, total columns
        sheet.Cells.Merge(0, 0, 1, 3); // A1:C1

        // Merge cells A2:A4 so the label spans three rows
        sheet.Cells["A2"].PutValue("Group");
        sheet.Cells.Merge(1, 0, 3, 1); // A2:A4

        // Optional: center the text in the merged cells
        Style mergedStyle = sheet.Cells["A1"].GetStyle();
        mergedStyle.HorizontalAlignment = TextAlignmentType.Center;
        mergedStyle.VerticalAlignment = TextAlignmentType.Center;
        sheet.Cells["A1"].SetStyle(mergedStyle);
        sheet.Cells["A2"].SetStyle(mergedStyle);

        // Configure PDF save options to keep the layout intact
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Prevent automatic page breaks that could split merged cells
            OnePagePerSheet = false,
            // Force all columns of a sheet onto a single PDF page to preserve spanning
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as PDF; merged cells will retain their spanning
        workbook.Save("MergedCells.pdf", pdfOptions);
    }
}
