// Title: Bold Text via HTML in Aspose.Cells Smart Marker (C#)
// Description: Demonstrates how to assign an HtmlString containing a <b> tag to a smart marker, bind a DataTable as the data source, process the marker with WorkbookDesigner, and save the workbook as HTML so the merged value appears in bold.
// Keywords: Aspose.Cells | smart marker | HtmlString | bold text | C# | WorkbookDesigner | HTML export | Excel formatting | data binding | dynamic report
// Common Searches: Aspose.Cells set smart marker bold | HtmlString smart marker C# example | how to make smart marker output bold | WorkbookDesigner process HTML tags | export Excel to HTML with bold smart markers
// Developer Intent: Apply HTML formatting to a smart marker so the merged cell displays bold text after processing.
// Use Cases: Generate HTML reports where specific fields (e.g., names) are emphasized in bold. | Create Excel templates that retain bold styling when exported to HTML via smart markers. | Automate dynamic documents with bold headings or labels using data‑driven smart markers.
// AI Prompts: Show how to use HtmlString with a smart marker to render bold text in Aspose.Cells for .NET. | Provide a C# example that processes a <b>${Field}</b> smart marker using WorkbookDesigner and saves as HTML. | Explain how to combine HTML tags with smart markers to apply styles such as bold, italic, or color.

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates how to assign an HtmlString containing a <b> tag to a smart marker, bind a DataTable as the data source, process the marker with WorkbookDesigner, and save the workbook as HTML so the merged value appears in bold.
class SmartMarkerBoldDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Place a smart marker inside an HTML string.
            // The <b> tag will make the value rendered in bold after processing.
            worksheet.Cells["A1"].HtmlString = "<b>${Name}</b>";

            // Prepare a data source for the smart marker
            DataTable table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("John Doe");

            // Process the smart marker using WorkbookDesigner (correct API)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(table);
            designer.Process();

            // Save the result as an HTML file to see the bold rendering
            workbook.Save("SmartMarkerBold.html", SaveFormat.Html);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
