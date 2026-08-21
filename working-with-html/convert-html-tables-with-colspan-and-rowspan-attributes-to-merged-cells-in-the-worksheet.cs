// Title: Convert HTML table colspan/rowspan to merged cells in Excel using Aspose.Cells for .NET
// Description: Demonstrates how to load an HTML file containing <table> elements with colspan or rowspan attributes into an Aspose.Cells Workbook, automatically converting those attributes into merged cells, and then saving the result as an XLSX file.
// Keywords: Aspose.Cells HTML to Excel | colspan to merged cells | rowspan to merged cells | HTML table import .NET | HtmlLoadOptions merged cells | convert HTML tables Excel
// Common Searches: Aspose.Cells preserve colspan when importing HTML | how to merge cells from HTML table in Excel .NET | load HTML with rowspan into Excel using Aspose | convert web table to Excel merged cells | HTML to XLSX merged cells Aspose.Cells
// Developer Intent: Load an HTML document that contains tables with colspan/rowspan and export it to an Excel workbook where the original merged‑cell layout is retained.
// Use Cases: Transform web‑based reports into Excel while keeping complex table structures. | Automate conversion of HTML email newsletters into spreadsheets with proper merged cells. | Generate Excel files from HTML templates that use colspan/rowspan for layout design.
// AI Prompts: Write C# code that uses Aspose.Cells to import an HTML file with colspan and rowspan and saves it as an XLSX preserving merged cells. | Explain the role of HtmlLoadOptions in handling colspan/rowspan during HTML to Excel conversion with Aspose.Cells. | Show how to programmatically verify that merged cells were created after loading an HTML table in Aspose.Cells.

using Aspose.Cells;

// Demonstrates how to load an HTML file containing <table> elements with colspan or rowspan attributes into an Aspose.Cells Workbook, automatically converting those attributes into merged cells, and then saving the result as an XLSX file.
class HtmlTableToMergedCells
{
    static void Main()
    {
        // Create an empty workbook (required by the lifecycle rule)
        Workbook workbook = new Workbook();

        // Prepare HTML load options (default settings are sufficient)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // Load the HTML file. Aspose.Cells automatically converts <table> elements
        // with colspan/rowspan attributes into merged cells in the worksheet.
        workbook = new Workbook("input.html", loadOptions);

        // Save the workbook to an Excel file. The merged cells reflect the original HTML layout.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
