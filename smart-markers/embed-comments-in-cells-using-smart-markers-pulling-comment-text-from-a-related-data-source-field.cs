// Title: C# – Insert dynamic Excel comments from a DataTable with Aspose.Cells smart markers
// Description: Shows how to bind a DataTable (Product, Comment) to a workbook, place a smart marker in a cell, embed a smart marker inside a comment note, define the required named range, run WorkbookDesigner.Process, and save the file with row‑wise comments.
// Keywords: Aspose.Cells | C# smart markers | Excel comments | WorkbookDesigner | DataTable binding | dynamic comments | named range _CellsSmartMarkers | populate comments from database | automated Excel reporting
// Common Searches: Aspose.Cells add comment using smart marker C# | How to bind DataTable to smart markers for Excel comments | WorkbookDesigner generate row comments from table | Define _CellsSmartMarkers range for comment smart markers | C# example smart marker comments in Excel
// Developer Intent: Create Excel comments automatically by mapping a data‑source field to a comment smart marker.
// Use Cases: Product catalog where each item’s description appears as a cell comment. | Audit report that attaches reviewer remarks to specific cells. | Financial spreadsheet that adds footnote text from a remarks table. | Customer‑feedback sheet that pulls comments from a CRM export.
// AI Prompts: Provide C# code that uses Aspose.Cells smart markers to fill cell comments from a DataTable without loops. | Explain why a named range called _CellsSmartMarkers is required for comment smart markers and how to set it. | Show how to retrieve and modify comment smart markers after WorkbookDesigner.Process.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkerComments
{
    // Shows how to bind a DataTable (Product, Comment) to a workbook, place a smart marker in a cell, embed a smart marker inside a comment note, define the required named range, run WorkbookDesigner.Process, and save the file with row‑wise comments.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a data source (DataTable) that contains the comment text
                DataTable dt = new DataTable("Products");
                dt.Columns.Add("Product", typeof(string));
                dt.Columns.Add("Comment", typeof(string));

                dt.Rows.Add("Apple", "Fresh and juicy");
                dt.Rows.Add("Banana", "Rich in potassium");
                dt.Rows.Add("Cherry", "Small but sweet");

                // Insert product names into column A using a smart marker
                sheet.Cells["A2"].PutValue("&=$Product");

                // Add a comment to cell B2; the comment will be duplicated for other rows by the designer
                int commentIndex = sheet.Comments.Add("B2");
                Comment comment = sheet.Comments[commentIndex];
                // Use a smart marker inside the comment note; it will be replaced with the value from the data source
                comment.Note = "&=$Comment";

                // Define a named range that contains the smart markers (required for the designer)
                AsposeRange dataRange = sheet.Cells.CreateRange("A2:B2");
                dataRange.Name = "_CellsSmartMarkers";

                // Set the data source for the designer
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);

                // Process the smart markers – this will populate the cells and replace the comment smart marker
                designer.Process();

                // Save the workbook
                string outputPath = "SmartMarkerComments.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
