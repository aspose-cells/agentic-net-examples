// Title: Convert an HTML table containing formulas to an Excel workbook and enable automatic calculation with Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML string via HtmlLoadOptions, loads it into a Workbook, sets the calculation mode to Automatic, evaluates all formulas, and saves the result as an .xlsx file. | Show how to import an HTML stream into an Aspose.Cells Workbook, trigger workbook.CalculateFormula, and export the workbook with automatic formula recalculation enabled. | Provide a snippet that demonstrates converting an HTML table with embedded Excel formulas into a Workbook, configuring automatic calculation, and saving to SaveFormat.Xlsx.
// Common Searches: aspnet convert html table with formulas to xlsx using aspose.cells | c# load html string into workbook and recalculate formulas automatically | aspose.cells set calculation mode automatic after loading html | how to evaluate formulas in workbook after importing html with aspose.cells | aspose.cells html to excel conversion with formula evaluation
// Tags: html-to-workbook conversion Aspose.Cells | automatic formula calculation Aspose.Cells | load html stream Aspose.Cells | save workbook xlsx Aspose.Cells | evaluate embedded formulas Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// // Converts an HTML string containing a table with formulas into an Excel workbook, triggers automatic formula calculation, and saves the file as .xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // HTML content to be converted to Excel
            string html = "<html><body><table>" +
                          "<tr><td>10</td><td>20</td></tr>" +
                          "<tr><td>=A1+B1</td><td>30</td></tr>" +
                          "</table></body></html>";

            // Convert HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream ms = new MemoryStream(htmlBytes))
            {
                // Load workbook from HTML stream
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(ms, loadOptions);

                // Calculate all formulas
                workbook.CalculateFormula();

                // Save the workbook as an Excel file
                string outputPath = "ConvertedFromHtml.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
