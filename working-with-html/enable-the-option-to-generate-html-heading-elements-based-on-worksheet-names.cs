using System;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with multiple worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sales");
            workbook.Worksheets.Add("Inventory");

            // Add some sample data to each sheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue($"Data for {sheet.Name}");
            }

            // Set a page header that contains the worksheet name (&A placeholder)
            // This will be rendered as an HTML heading when ExportPageHeaders is true
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Center section (1) will display the sheet name
                sheet.PageSetup.SetHeader(1, "&A");
            }

            // Configure HTML save options to export page headers
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export the page header (which now contains the sheet name)
                ExportPageHeaders = true,
                // Save all sheets into a single HTML file for easier viewing
                SaveAsSingleFile = true,
                // Optional: give the HTML page a title
                PageTitle = "Workbook with Sheet Name Headings"
            };

            // Save the workbook as HTML
            workbook.Save("WorkbookWithSheetHeadings.html", htmlOptions);
        }
    }
}