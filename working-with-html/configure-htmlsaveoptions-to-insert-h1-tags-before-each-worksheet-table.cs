using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Header1");
            sheet1.Cells["B1"].PutValue("Header2");
            sheet1.Cells["A2"].PutValue("Data1");
            sheet1.Cells["B2"].PutValue("Data2");

            // Add a second worksheet for demonstration
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheetIndex];
            sheet2.Name = "SecondSheet";
            sheet2.Cells["A1"].PutValue("ColA");
            sheet2.Cells["B1"].PutValue("ColB");
            sheet2.Cells["A2"].PutValue("Val1");
            sheet2.Cells["B2"].PutValue("Val2");

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Export row/column headings if needed
                ExportRowColumnHeadings = true,

                // NOTE: Aspose.Cells does not provide a direct property to inject
                // custom <h1> tags before each worksheet's HTML table.
                // To achieve this, one would need to post‑process the generated HTML
                // (e.g., read the file, insert <h1>{WorksheetName}</h1> before each <table>,
                // and then write the modified content back).
                // The placeholder below marks where such custom logic would be applied.
                // InsertCustomHeadings = true; // <-- No such API exists (placeholder)
            };

            // Save the workbook as a single HTML file (contains all worksheets)
            workbook.Save("output.html", options);

            // -----------------------------------------------------------------
            // Post‑processing placeholder (not part of Aspose.Cells API):
            // -----------------------------------------------------------------
            // string html = System.IO.File.ReadAllText("output.html");
            // // Insert <h1> tags before each <table> based on worksheet names...
            // System.IO.File.WriteAllText("output.html", html);
            // -----------------------------------------------------------------
        }
    }
}