// Title: Export Worksheet Headings as Separate HTML Tables with Aspose.Cells for .NET
// Description: Demonstrates how to enable HtmlSaveOptions.ExportHeadings in Aspose.Cells (C#) so that row and column headings are saved as distinct HTML tables when exporting a workbook.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportHeadings | C# HTML export | worksheet headings HTML | separate tables | Excel to HTML | Aspose.Cells .NET
// Common Searches: Aspose.Cells export headings to HTML | HtmlSaveOptions ExportHeadings C# | Save Excel worksheet headings as separate HTML tables | How to generate HTML with column and row headings using Aspose.Cells | Export row/column headers as distinct tables Aspose.Cells .NET
// Developer Intent: Save a workbook to HTML with row and column headings emitted as independent tables.
// Use Cases: Web reports that require header rows/columns to be styled separately from data. | Accessible HTML spreadsheets where headings are isolated for screen‑reader navigation. | Printable HTML output with distinct heading sections for custom page layout. | Embedding Excel data in web pages while applying CSS only to header tables.
// AI Prompts: Provide a C# example that saves an Aspose.Cells workbook to HTML with ExportHeadings enabled. | Show how to export only the worksheet’s row and column headings as separate HTML tables using Aspose.Cells. | What is the recommended modern alternative to the obsolete ExportHeadings property for exporting headings in Aspose.Cells?

using System;
using Aspose.Cells;

namespace ExportHeadingsExample
{
    // Demonstrates how to enable HtmlSaveOptions.ExportHeadings in Aspose.Cells (C#) so that row and column headings are saved as distinct HTML tables when exporting a workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");
            sheet.Cells["A3"].PutValue("Data3");
            sheet.Cells["B3"].PutValue("Data4");

            // Configure HTML save options to export row/column headings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // This property is obsolete but still works for the requested scenario
            saveOptions.ExportHeadings = true;

            // Save the workbook as HTML; headings will appear as separate tables
            string outputPath = "WorksheetHeadings.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath} with headings exported as separate HTML tables.");
        }
    }
}
