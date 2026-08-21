// Title: C# – Insert Dynamic Hyperlinks into Excel Using Aspose.Cells Smart Markers
// Description: This example creates a workbook, defines smart markers for product names and URLs, fills a DataTable with dynamic links (e.g., Aspose.Cells, GitHub, Stack Overflow), processes the markers with **WorkbookDesigner**, then adds clickable hyperlinks to the product‑name cells and saves the file as **ProductsWithHyperlinks.xlsx**. Learn how to generate Excel reports with live web links in a single pass.
// Keywords: Aspose.Cells C# hyperlink smart markers | add Excel hyperlink programmatically | dynamic URL Excel export .NET | WorkbookDesigner smart marker example | C# generate product catalog Excel | Aspose.Cells GitHub sample | Excel hyperlink from DataTable | smart markers external links
// Common Searches: how to add hyperlinks with smart markers in Aspose.Cells | Aspose.Cells C# create Excel file with clickable URLs | populate Excel hyperlinks from a database using WorkbookDesigner | dynamic hyperlink generation Aspose.Cells .NET | smart marker syntax for URLs Aspose.Cells
// Developer Intent: Generate an Excel workbook where each product name cell links to its website by using smart markers and a DataTable of URLs.
// Use Cases: Automated product catalog where every item links to the vendor site. | Exporting a list of resources (e.g., documentation, repositories) to Excel with one‑click access. | Creating a sales report that includes live hyperlinks to order details or contracts.
// AI Prompts: Show how to apply blue‑underline hyperlink styling to the linked cells after they are added. | Rewrite the sample to use a List<T> as the data source instead of a DataTable while preserving smart‑marker hyperlink insertion. | Explain how to add hyperlinks to an entire column in a single call using Hyperlinks.Add with a range and dynamic URLs.

using System;
using System.Data;
using Aspose.Cells;

// This example creates a workbook, defines smart markers for product names and URLs, fills a DataTable with dynamic links (e.g., Aspose.Cells, GitHub, Stack Overflow), processes the markers with **WorkbookDesigner**, then adds clickable hyperlinks to the product‑name cells and saves the file as **ProductsWithHyperlinks.xlsx**. Learn how to generate Excel reports with live web links in a single pass.
class InsertHyperlinksViaSmartMarkers
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Website");

            // Insert smart markers for product name and URL (starting from row 2)
            cells["A2"].PutValue("&=$ProductName");
            cells["B2"].PutValue("&=$ProductUrl");

            // Create a DataTable that will serve as the data source
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ProductName", typeof(string));
            dt.Columns.Add("ProductUrl", typeof(string));

            // Populate the table with dynamic URLs
            dt.Rows.Add("Aspose.Cells", "https://www.aspose.com/cells");
            dt.Rows.Add("GitHub", "https://github.com");
            dt.Rows.Add("Stack Overflow", "https://stackoverflow.com");

            // Set up the WorkbookDesigner with the workbook and the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource(dt);

            // Process the smart markers (placeholders will be replaced with actual values)
            designer.Process();

            // Add hyperlinks to the product name cells using URLs from column B
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int rowIndex = i + 1; // Data starts at row 2 in the sheet (zero‑based index)
                string url = dt.Rows[i]["ProductUrl"].ToString();

                // Add a hyperlink to the cell in column A (product name) that points to the URL
                sheet.Hyperlinks.Add(rowIndex, 0, 1, 1, url);
            }

            // Save the workbook to an Excel file
            workbook.Save("ProductsWithHyperlinks.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
