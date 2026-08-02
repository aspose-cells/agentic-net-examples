// Title: Export Workbook to HTML with WidthScalable (em‑based column sizing) in C#
// Description: Shows how to create a workbook, add sample data, enable HtmlSaveOptions.WidthScalable to use em units for column widths, and save the result as responsive HTML with Aspose.Cells for .NET.
// Keywords: Aspose.Cells HtmlSaveOptions WidthScalable | C# export Excel to HTML | em based column width | responsive HTML from Excel | Aspose.Cells HTML scaling | save workbook as HTML .NET | column width scaling
// Common Searches: Aspose.Cells WidthScalable example C# | how to enable em column sizing when saving Excel to HTML | export Excel to responsive HTML using Aspose.Cells | HtmlSaveOptions WidthScalable property usage | C# generate HTML with scalable column widths from workbook
// Developer Intent: The developer needs to export an Excel workbook to HTML where column widths are defined in em units by setting HtmlSaveOptions.WidthScalable to true.
// Use Cases: Create web‑friendly reports that adapt column widths to font size on different devices. | Generate HTML email templates from spreadsheets without fixed pixel widths. | Build responsive dashboards where column layout scales with surrounding CSS.
// AI Prompts: Provide a C# snippet that saves an Aspose.Cells workbook to HTML with WidthScalable enabled and custom CSS. | Explain the impact of the WidthScalable property on column width units and how to fine‑tune the scaling factor. | Give a step‑by‑step guide to convert an Excel file to responsive HTML using HtmlSaveOptions.WidthScalable in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsWidthScalableDemo
{
    // Shows how to create a workbook, add sample data, enable HtmlSaveOptions.WidthScalable to use em units for column widths, and save the result as responsive HTML with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Column A");
            sheet.Cells["B1"].PutValue("Column B");
            sheet.Cells["A2"].PutValue("Short");
            sheet.Cells["B2"].PutValue("This is a longer text that will demonstrate column width scaling.");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable em‑based column sizing (WidthScalable = true)
            htmlOptions.WidthScalable = true;

            // Save the workbook as HTML (lifecycle rule: save)
            workbook.Save("WidthScalableOutput.html", htmlOptions);
        }
    }
}
