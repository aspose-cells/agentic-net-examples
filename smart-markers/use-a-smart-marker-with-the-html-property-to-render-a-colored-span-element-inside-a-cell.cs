// Title: Color Text with an HTML <span> Smart Marker in Aspose.Cells for .NET
// Description: Demonstrates how to set a cell's HtmlString to an HTML <span> that contains a smart marker, apply an inline red color style, bind a DataTable as the data source, process the marker with WorkbookDesigner, and save the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | smart marker | HtmlString | HTML span | inline CSS | colored text in Excel | WorkbookDesigner | DataTable binding | Excel report styling
// Common Searches: Aspose.Cells smart marker inside HTML span | set cell HtmlString with smart marker C# | apply inline CSS to smart marker Aspose.Cells | color text using smart markers in .NET | WorkbookDesigner process HTML smart markers
// Developer Intent: Generate an Excel workbook where a smart marker embedded in an HTML <span> renders data with red-colored text.
// Use Cases: Create reports that highlight dynamic values (e.g., employee names) in a specific color without using Excel conditional formatting. | Produce templated Excel files where HTML styling is applied directly through smart markers. | Integrate data‑driven styling in automated spreadsheet generation pipelines.
// AI Prompts: Show how to embed a smart marker inside an HTML <span> with a custom text color using Aspose.Cells for .NET. | Provide a C# example that uses WorkbookDesigner to process HTML smart markers containing inline CSS. | Explain the steps to bind a DataTable to a smart marker placed in a cell's HtmlString and render colored text in the output workbook.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHtml
{
    // Demonstrates how to set a cell's HtmlString to an HTML <span> that contains a smart marker, apply an inline red color style, bind a DataTable as the data source, process the marker with WorkbookDesigner, and save the workbook as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set a smart marker inside an HTML span element.
                // The smart marker &=Employees.Name& will be replaced with the value from the data source.
                // The span uses an inline style to set the text color to red.
                sheet.Cells["A1"].HtmlString = "<span style='color:#FF0000'>&=Employees.Name&</span>";

                // Prepare a data source (DataTable) that matches the smart marker name.
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Name", typeof(string));
                dt.Rows.Add("John Doe");
                dt.Rows.Add("Jane Smith");

                // Process the smart markers using WorkbookDesigner.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Save the workbook.
                workbook.Save("SmartMarkerHtmlSpan.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
