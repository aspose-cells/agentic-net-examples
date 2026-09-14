// Title: Create clickable Excel hyperlinks from DataTable values using Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that employs Aspose.Cells WorkbookDesigner to replace smart markers with URL and display text taken from a DataTable, then adds a Hyperlink object to each display cell. | Demonstrate how to iterate over the rows after smart marker processing and attach a hyperlink to the text cell using the URL stored in the adjacent column. | Provide a full example that saves the workbook as an .xlsx file with all generated hyperlinks functional.
// Common Searches: aspnet c# Aspose.Cells add hyperlink to cell after smart marker processing | how to bind a DataTable to smart markers and create clickable links in Excel using Aspose.Cells | C# example for inserting dynamic URLs into Excel with smart markers and Hyperlinks class
// Tags: Aspose.Cells WorkbookDesigner smart marker population | Aspose.Cells add hyperlink to Excel cell | C# generate Excel hyperlinks from DataTable | smart markers dynamic URL insertion | Excel .xlsx hyperlink creation with Aspose

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHyperlinkDemo
{
    // Shows how to use Aspose.Cells WorkbookDesigner to fill smart markers with URL and display text from a DataTable, then adds Hyperlink objects to the corresponding cells for each row and saves the workbook as an .xlsx file with functional clickable links.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert smart markers for URL and display text
            // A1 will hold the URL, B1 will hold the text to display
            sheet.Cells["A1"].PutValue("&=LinkUrl");
            sheet.Cells["B1"].PutValue("&=LinkText");

            // Prepare a data source with dynamic URL fields
            DataTable dt = new DataTable("Links");
            dt.Columns.Add("LinkUrl", typeof(string));
            dt.Columns.Add("LinkText", typeof(string));

            // Add sample rows (in real scenarios this could come from a database, JSON, etc.)
            dt.Rows.Add("https://www.aspose.com", "Aspose Home");
            dt.Rows.Add("https://github.com/aspose-cells", "Aspose Cells GitHub");
            dt.Rows.Add("https://www.google.com", "Google");

            // Use WorkbookDesigner to process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // Populate A1 and B1 with the first row values

            // After processing, the first row values are in A1 (URL) and B1 (display text)
            // Add a hyperlink to the display text cell (B1) using the URL from A1
            string url = sheet.Cells["A1"].StringValue; // Get the populated URL
            sheet.Hyperlinks.Add("B1", 1, 1, url); // Create hyperlink on B1

            // Optionally, repeat for remaining rows by inserting them below
            // Insert remaining rows starting from row 2
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                int rowIndex = i + 1; // Excel rows are 1‑based; row 2 corresponds to index 1
                // Populate URL and display text
                sheet.Cells[rowIndex, 0].PutValue(dt.Rows[i]["LinkUrl"]);
                sheet.Cells[rowIndex, 1].PutValue(dt.Rows[i]["LinkText"]);

                // Add hyperlink to the display text cell
                sheet.Hyperlinks.Add(rowIndex, 1, 1, 1, dt.Rows[i]["LinkUrl"].ToString());
            }

            // Save the workbook to an Excel file
            workbook.Save("SmartMarkerHyperlinksDemo.xlsx");
        }
    }
}
