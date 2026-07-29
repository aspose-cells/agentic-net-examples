// Title: C# – Delete Blank Columns and Export Trimmed Worksheet to PDF with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells for .NET, removes every column that contains no data from the first worksheet with DeleteBlankColumns, and saves the cleaned sheet directly as a PDF file.
// Keywords: Aspose.Cells DeleteBlankColumns | C# Excel to PDF | remove empty columns Aspose.Cells | trim worksheet before PDF export | .NET export Excel as PDF | clean Excel report C#
// Common Searches: Aspose.Cells delete blank columns C# | how to export trimmed Excel sheet to PDF using Aspose.Cells | remove empty columns from worksheet and save as PDF .NET | DeleteBlankColumns example Aspose.Cells | convert Excel to PDF after deleting blank columns
// Developer Intent: Eliminate all empty columns from an Excel worksheet and generate a PDF of the resulting compact sheet.
// Use Cases: Prepare printable reports by stripping unused columns before PDF conversion. | Automate batch processing of workbooks where each sheet must be compacted for archiving. | Reduce file size and visual whitespace in PDFs generated from data extracts.
// AI Prompts: Generate C# code that iterates through every worksheet in a workbook, deletes blank columns on each sheet, and saves the workbook as a PDF using Aspose.Cells. | Explain how Worksheet.Cells.DeleteBlankColumns determines a column is blank and what options affect its behavior. | Show a combined example that deletes both blank rows and columns before exporting the worksheet to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTrimAndExportPdf
{
    // Loads an Excel workbook using Aspose.Cells for .NET, removes every column that contains no data from the first worksheet with DeleteBlankColumns, and saves the cleaned sheet directly as a PDF file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Path for the resulting PDF file
            string pdfFile = "output.pdf";

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(sourceFile);

            // Access the first worksheet (you can iterate if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Delete all blank columns in the worksheet
            worksheet.Cells.DeleteBlankColumns();

            // Save the trimmed worksheet as a PDF document
            workbook.Save(pdfFile, SaveFormat.Pdf);
        }
    }
}
