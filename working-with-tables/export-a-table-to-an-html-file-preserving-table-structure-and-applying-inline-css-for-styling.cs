// Title: Export an Excel worksheet table to HTML with inline CSS using Aspose.Cells for .NET
// Common Searches: how to save an Excel ListObject as an HTML file with inline styles using Aspose.Cells C# | Aspose.Cells export worksheet table to HTML preserving table layout | C# generate HTML from Excel range with embedded CSS using Aspose.Cells | save Excel workbook as HTML with table formatting and inline CSS in .NET | Aspose.Cells HtmlSaveOptions example for exporting styled tables
// Tags: Aspose.Cells HtmlSaveOptions for table export | ListObject HTML conversion with inline styling | maintain Excel table layout in generated HTML | C# apply header row style before HTML save | create output folder for Aspose.Cells HTML save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills a range with product data, styles the header row, converts the range into a ListObject to keep table semantics, configures HtmlSaveOptions, ensures the target directory exists, and saves the workbook as an HTML file where the table structure and inline CSS are preserved.
class ExportTableToHtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (a simple table).
            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            // Data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(60);
            sheet.Cells["C4"].PutValue(1.2);

            // Apply simple styling to the header row.
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells.CreateRange("A1:C1").ApplyStyle(headerStyle, flag);

            // Convert the range into a table (ListObject) to preserve structure in HTML.
            int totalRows = 4; // including header
            int totalCols = 3;
            int firstRow = 0, firstColumn = 0, lastRow = totalRows - 1, lastColumn = totalCols - 1;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "ProductTable";
            table.ShowHeaderRow = true;
            // Correct property name for hiding the total row.
            table.ShowTotals = false;

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Determine output path and ensure the directory exists.
            string outputPath = "ExportedTable.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
                outputPath = Path.Combine(outputDir, outputPath);
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file.
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
