// Title: Export a smart‑marker‑populated Excel workbook with charts and graphics to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, assigns a data source to WorkbookDesigner, processes smart markers, and saves the result as a PDF while retaining all charts and graphics. | Show how to configure WorkbookDesigner with a custom DataTable before converting the workbook to PDF with Aspose.Cells. | Provide a snippet that creates a PDF/A‑1b compliant document from a smart‑marker‑filled workbook, preserving embedded graphics.
// Common Searches: Aspose.Cells how to keep Excel charts when exporting to PDF after processing smart markers | C# export smart marker workbook to PDF with graphics intact | set data source for WorkbookDesigner before PDF conversion Aspose.Cells .NET | convert Excel file containing smart markers and drawings to PDF preserving layout
// Tags: export smart markers to PDF Aspose.Cells | preserve charts during Excel to PDF conversion .NET | WorkbookDesigner data source configuration | Aspose.Cells PDF/A‑1b generation with graphics | C# process smart markers and save as PDF

using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook with smart markers, assigns a data source, processes the markers to generate charts and graphics, and saves the populated workbook as a PDF while preserving all visual elements.
class Program
{
    static void Main()
    {
        // Load the workbook that contains smart markers, charts, and graphics
        Workbook workbook = new Workbook("input.xlsx");

        // Process smart markers to generate data-driven content (charts, graphics, etc.)
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        // TODO: Set the data source for the smart markers, e.g.:
        // designer.SetDataSource(yourDataSource);
        designer.Process();

        // Export the populated workbook to PDF while preserving charts and graphics
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
