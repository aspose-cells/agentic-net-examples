using System;
using System.Data;
using System.IO;
using Aspose.Cells;

class InsertHyperlinksViaSmartMarkers
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Prepare the template with smart markers
            //    Column A will hold the display text, Column B will hold the URL
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Link");

            // Smart markers: &=$Data.DisplayText and &=$Data.Url
            sheet.Cells["A2"].PutValue("&=$Data.DisplayText");
            sheet.Cells["B2"].PutValue("&=$Data.Url");

            // 3. Create a data source (DataTable) with dynamic URL fields
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("DisplayText", typeof(string));
            dt.Columns.Add("Url", typeof(string));

            dt.Rows.Add("Aspose Home", "https://www.aspose.com");
            dt.Rows.Add("GitHub", "https://github.com");
            dt.Rows.Add("Stack Overflow", "https://stackoverflow.com");

            // 4. Set up WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource("Data", dt);

            // 5. Process the smart markers – this populates the cells with display text and URLs
            designer.Process();

            // 6. After processing, add hyperlinks to the display cells using the URLs from column B
            //    Loop through the populated rows (starting from row 2)
            for (int row = 1; row <= dt.Rows.Count; row++)
            {
                // Cell with display text (Column A)
                Cell displayCell = sheet.Cells[row, 0];

                // Corresponding URL cell (Column B)
                string url = sheet.Cells[row, 1].StringValue;

                // Add hyperlink to the display cell
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
                sheet.Hyperlinks.Add(displayCell.Row, displayCell.Column, 1, 1, url);

                // Optionally, clear the URL cell if you don't want it visible
                sheet.Cells[row, 1].PutValue(string.Empty);
            }

            // 7. Save the workbook (using the standard save lifecycle)
            string outputPath = "HyperlinksViaSmartMarkers.xlsx";

            // Ensure we can write to the path (no need for File.Exists check when creating a new file)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}