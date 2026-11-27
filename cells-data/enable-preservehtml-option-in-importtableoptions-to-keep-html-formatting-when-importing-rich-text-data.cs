// Title: How to preserve HTML formatting when importing a DataTable into an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import DataTable with HTML tags retained in cells | set cell HtmlString property to keep bold and italic formatting in Excel workbook | preserve rich text formatting from DataTable when saving as .xlsx with Aspose.Cells | C# example of writing HTML‑styled text to Excel cells using Aspose.Cells | avoid losing HTML markup during DataTable to Excel conversion with Aspose.Cells
// Tags: DataTable to Excel HTML formatting Aspose.Cells | cell.HtmlString usage .NET | preserve rich text during Excel export C# | Aspose.Cells HTML string rendering | export bold italic text to .xlsx with Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The sample creates a workbook, builds a DataTable containing HTML‑formatted strings, assigns each cell's HtmlString property with the HTML content to retain bold and italic styling, and saves the workbook as PreserveHtmlOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a DataTable containing HTML-formatted text
            DataTable dt = new DataTable();
            dt.Columns.Add("HtmlColumn", typeof(string));
            DataRow row = dt.NewRow();
            row["HtmlColumn"] = "<b>Bold Text</b> and <i>Italic Text</i>";
            dt.Rows.Add(row);

            // Import the DataTable while preserving HTML formatting
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    Cell cell = sheet.Cells[i, j];
                    // Set the cell's HTML string; Aspose.Cells will render it appropriately
                    cell.HtmlString = dt.Rows[i][j].ToString();
                }
            }

            // Save the workbook
            string outputPath = "PreserveHtmlOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
