using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for the worksheet
        worksheet.DisplayRightToLeft = true;

        // Add some RTL (e.g., Hebrew) text to a cell
        worksheet.Cells["A1"].PutValue("שלום עולם"); // "Hello World" in Hebrew

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the workbook to an HTML file
        workbook.Save("RtlExported.html", htmlOptions);
    }
}