// Title: Configure custom page header and footer text and freeze the first three rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set left, center, right page header text | How to freeze the first three rows after adding a header with Aspose.Cells | Repeat header rows on every printed page using PrintTitleRows in Aspose.Cells .NET | Set footer with page number and current time in Aspose.Cells C# example | Aspose.Cells example for combining PrintTitleRows and FreezePanes
// Tags: page header/footer configuration Aspose.Cells | freeze panes for header rows Aspose.Cells | PrintTitleRows usage Aspose.Cells | header/footer formatting with fonts Aspose.Cells | save workbook to XLSX format Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, adds three header rows, configures left, center, and right sections of the page header and footer via PageSetup, repeats those rows on each printed page using PrintTitleRows, freezes rows 1‑3 with FreezePanes, and saves the result as HeaderFooterFreezeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Report";

            // ------------------------------------------------------------
            // 2. Populate the first three rows with header information.
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Company XYZ");
            worksheet.Cells["A2"].PutValue("Sales Report");
            worksheet.Cells["A3"].PutValue(DateTime.Now.ToString("yyyy-MM-dd"));

            // ------------------------------------------------------------
            // 3. Configure page header and footer text.
            //    Header sections: 0 = left, 1 = center, 2 = right.
            //    Footer sections: same indexing.
            // ------------------------------------------------------------
            // Header
            worksheet.PageSetup.SetHeader(0, "&\"Arial,Bold\"&12 Company XYZ"); // left
            worksheet.PageSetup.SetHeader(1, "&\"Arial,Bold\"&12 Sales Report"); // center
            worksheet.PageSetup.SetHeader(2, "&D"); // right – current date

            // Footer
            worksheet.PageSetup.SetFooter(0, "Confidential"); // left
            worksheet.PageSetup.SetFooter(1, "Page &P of &N"); // center – page number of total pages
            worksheet.PageSetup.SetFooter(2, "&T"); // right – current time

            // Optional: repeat the header rows on every printed page.
            worksheet.PageSetup.PrintTitleRows = "1:3";

            // ------------------------------------------------------------
            // 4. Freeze the rows that contain the header (rows 1‑3).
            //    FreezePanes(string cellName, int freezedRows, int freezedColumns)
            //    freezes rows above the specified cell.
            // ------------------------------------------------------------
            worksheet.FreezePanes("A4", 3, 0); // freezes rows 1‑3

            // ------------------------------------------------------------
            // 5. Save the workbook.
            // ------------------------------------------------------------
            const string outputPath = "HeaderFooterFreezeDemo.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
