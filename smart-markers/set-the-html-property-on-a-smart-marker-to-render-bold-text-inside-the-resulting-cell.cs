// Title: Bold Text in Smart Markers using HtmlString – Aspose.Cells for .NET (C#)
// Description: Shows how to embed an HTML <b> tag in a smart marker (e.g., "<b>&=Name</b>") so the merged cell renders the value in bold after WorkbookDesigner processing. The example creates a workbook, sets the HtmlString on cell A1, binds a DataTable with a Name column, processes the markers, and saves the result as HTML.
// Keywords: Aspose.Cells | smart markers | HtmlString | bold text | C# | .NET | WorkbookDesigner | HTML formatting | data binding | save as HTML
// Common Searches: Aspose.Cells bold smart marker C# | HtmlString smart marker example | apply HTML tags to smart markers .NET | render bold placeholder with WorkbookDesigner | save smart marker output as HTML Aspose.Cells
// Developer Intent: Add HTML markup to a smart marker so the merged cell displays the field value in bold after processing.
// Use Cases: Generate HTML reports where customer names appear bold using a smart marker. | Create invoice sheets that highlight product titles in bold by embedding <b> tags in smart markers. | Produce email‑template spreadsheets where placeholder fields are bold after smart‑marker processing.
// AI Prompts: Provide a C# example that uses HtmlString to apply multiple HTML styles (italic, color, underline) to smart markers in Aspose.Cells. | Show how to process smart markers with WorkbookDesigner and save the workbook as PDF while preserving the HTML bold formatting. | Explain how to bind a list of objects to smart markers and format each property with different HTML tags using Aspose.Cells for .NET.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // Shows how to embed an HTML <b> tag in a smart marker (e.g., "<b>&=Name</b>") so the merged cell renders the value in bold after WorkbookDesigner processing. The example creates a workbook, sets the HtmlString on cell A1, binds a DataTable with a Name column, processes the markers, and saves the result as HTML.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set a smart marker inside an HTML string to make the result bold
                worksheet.Cells["A1"].HtmlString = "<b>&=Name</b>";

                // Prepare a simple data source (DataTable) with a column matching the smart marker
                DataTable data = new DataTable();
                data.Columns.Add("Name", typeof(string));
                data.Rows.Add("John Doe");

                // Process the smart markers using WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(data);
                designer.Process();

                // Save the workbook (HTML format to see the bold rendering)
                string outputPath = "SmartMarkerBold.html";
                workbook.Save(outputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
