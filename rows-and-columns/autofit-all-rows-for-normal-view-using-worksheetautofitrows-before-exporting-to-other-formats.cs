// Title: C# – AutoFitRows in Aspose.Cells: Adjust all row heights before exporting to XLSX, PDF, PNG
// Description: Shows how to use Aspose.Cells for .NET to automatically fit every row in a worksheet, optionally auto‑fit columns, and then save the workbook as XLSX, PDF, or PNG. The example handles wrapped text and line breaks without manual height settings.
// Keywords: Aspose.Cells | C# | AutoFitRows | .NET | auto fit rows | export to PDF | export to PNG | Excel row height | Worksheet.AutoFitRows | auto fit columns
// Common Searches: Aspose.Cells auto fit rows C# | Worksheet.AutoFitRows example | How to adjust row height before PDF export Aspose.Cells | AutoFitRows before saving workbook | C# auto fit rows and columns Aspose.Cells
// Developer Intent: Automatically adjust the height of all rows in a worksheet before exporting the workbook to other formats.
// Use Cases: Generate PDF reports where wrapped text fits each row without manual sizing. | Create Excel files that retain proper row spacing when converted to PNG images for documentation. | Prepare XLSX workbooks with correctly sized rows for downstream processing or data import.
// AI Prompts: Provide C# code that calls Worksheet.AutoFitRows, then saves the workbook as PDF with custom page margins. | Show an example that auto‑fits only a specific range of rows in Aspose.Cells and exports the result to XLSX. | Explain how AutoFitRows handles cells containing line breaks and its effect on the layout of exported PDFs.

using System;
using Aspose.Cells;

// Shows how to use Aspose.Cells for .NET to automatically fit every row in a worksheet, optionally auto‑fit columns, and then save the workbook as XLSX, PDF, or PNG. The example handles wrapped text and line breaks without manual height settings.
class AutoFitRowsExample
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data that will require row height adjustment
        worksheet.Cells["A1"].PutValue("This is a test string for AutoFitRows demonstration");
        worksheet.Cells["A2"].PutValue("Another line of text\nwith line break to show row height adjustment");
        worksheet.Cells["B1"].PutValue("Column B content");

        // Auto‑fit all rows in the worksheet before any export
        worksheet.AutoFitRows();

        // (Optional) Auto‑fit columns for better visibility after rows are adjusted
        worksheet.AutoFitColumns();

        // Export the workbook to various formats after auto‑fitting
        workbook.Save("AutoFitRowsDemo.xlsx", SaveFormat.Xlsx);
        workbook.Save("AutoFitRowsDemo.pdf", SaveFormat.Pdf);
        workbook.Save("AutoFitRowsDemo.png", SaveFormat.Png);
    }
}
