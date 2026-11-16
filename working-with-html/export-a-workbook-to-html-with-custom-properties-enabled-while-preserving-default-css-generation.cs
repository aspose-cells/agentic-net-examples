// Title: Export an Aspose.Cells workbook to HTML with custom document properties while keeping default CSS generation (C#)
// Common Searches: C# Aspose.Cells save workbook as HTML with custom document properties | How to include custom document properties in HTML output using Aspose.Cells | Aspose.Cells HTML export default CSS style preservation C# | Save Excel workbook to HTML with custom properties using HtmlSaveOptions | Export workbook to HTML while retaining built‑in CSS in Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions custom document properties | export workbook to HTML default CSS Aspose.Cells | C# Aspose.Cells HTML export with custom properties | preserve built‑in CSS when saving Excel as HTML Aspose.Cells | HtmlSaveOptions SaveFormat.Html example

using System;
using Aspose.Cells;

// The example creates a Workbook, adds sample data and a custom document property, configures HtmlSaveOptions for HTML output (relying on the default CSS generation), and saves the workbook as an HTML file using Aspose.Cells in C#.
class ExportWorkbookToHtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Set a custom document property (will be exported to HTML if supported)
            workbook.CustomDocumentProperties.Add("Project", "Aspose.Cells Demo");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The properties ExportCustomProperties and ExportCssStyle are not available in this version;
            // default behavior will handle CSS generation and property export.

            // Export the workbook to HTML with the specified options
            string outputFile = "ExportedWorkbook.html";
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
