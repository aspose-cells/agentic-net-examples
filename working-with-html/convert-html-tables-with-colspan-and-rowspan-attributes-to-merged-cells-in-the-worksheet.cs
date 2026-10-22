// Title: How to convert an HTML table with colspan and rowspan into merged cells in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Load an HTML string that contains colspan and rowspan attributes into an Aspose.Cells Workbook and make the worksheet display the merged cells correctly. | Write C# code to read an HTML table with merged columns and rows, auto‑fit the worksheet columns, and save the result as an .xlsx file with Aspose.Cells. | Demonstrate configuring HtmlLoadOptions in a .NET application to retain cell merging when converting HTML to Excel.
// Common Searches: Aspose.Cells .NET load HTML table with colspan and rowspan preserving merged cells | C# convert HTML table to Excel with merged cells using HtmlLoadOptions | How to keep merged cells when importing HTML into an Aspose.Cells workbook | AutoFitColumns after loading HTML into Aspose.Cells workbook C#
// Tags: html to xlsx merged cells Aspose.Cells | load html with colspan Aspose.Cells C# | keep rowspan merged cells Aspose.Cells | autofit columns after html load Aspose.Cells | html load options preserve merging .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an HTML string that includes a table with colspan and rowspan attributes into an Aspose.Cells Workbook using HtmlLoadOptions, automatically adjusts column widths, and saves the worksheet as 'ConvertedTable.xlsx', ensuring that the original HTML merges are represented as merged cells in the Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // HTML string containing a table with colspan and rowspan attributes.
            string html = @"
<html>
<body>
<table border='1'>
    <tr>
        <th>Header 1</th>
        <th colspan='2'>Header 2-3</th>
    </tr>
    <tr>
        <td rowspan='2'>Row 1-2, Col 1</td>
        <td>R1C2</td>
        <td>R1C3</td>
    </tr>
    <tr>
        <td colspan='2'>R2C2-3 merged</td>
    </tr>
</table>
</body>
</html>";

            // Convert HTML string to a memory stream.
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (var htmlStream = new MemoryStream(htmlBytes))
            {
                // Load the HTML into a new workbook using HtmlLoadOptions.
                var loadOptions = new HtmlLoadOptions();
                var workbook = new Workbook(htmlStream, loadOptions);

                // Adjust column widths to fit the content.
                workbook.Worksheets[0].AutoFitColumns();

                // Define output file path.
                string outputPath = "ConvertedTable.xlsx";

                // Save the workbook to an Excel file.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
