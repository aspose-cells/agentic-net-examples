// Title: Set the HtmlString property on a smart‑marker result cell to display bold HTML text with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set HtmlString on a processed smart marker cell to show bold text | Render HTML tags inside an Excel cell after smart marker processing using Aspose.Cells | Display formatted bold text from a DataTable in a smart marker cell with Aspose.Cells | WorkbookDesigner HTML string output in Excel cell C# example | How to make a smart marker render HTML content in Aspose.Cells
// Tags: smart marker html rendering Aspose.Cells | set cell HtmlString property C# | WorkbookDesigner process html data source | excel cell display bold text from html | Aspose.Cells treat cell value as html

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, places a smart marker (<#=HtmlString#>) in cell A1, supplies a DataTable containing "<b>Bold Text</b>", processes the marker with WorkbookDesigner, then assigns the cell's HtmlString property to its string value so the HTML renders as bold text in Excel, calculates formulas, and saves the file as SmartMarkerBoldHtml.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Place a smart marker in cell A1 that will be replaced by the HTML string
            sheet.Cells["A1"].PutValue("<#=HtmlString#>");

            // Prepare the data source for the smart marker
            DataTable dt = new DataTable();
            dt.Columns.Add("HtmlString", typeof(string));
            dt.Rows.Add("<b>Bold Text</b>");

            // Process the smart marker using WorkbookDesigner (alternative to SmartMarkerProcessor)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // After processing, mark the resulting cell as HTML so it renders correctly
            Cell resultCell = sheet.Cells["A1"];
            resultCell.HtmlString = resultCell.StringValue; // treat the cell value as HTML

            // Optional: evaluate any formulas
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "SmartMarkerBoldHtml.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
