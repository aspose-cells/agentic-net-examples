// Title: Verify that <br> tags are converted to newline characters when importing HTML into an Excel worksheet cell using Aspose.Cells for .NET (C#)
// Common Searches: how to keep line breaks from <br> tags when loading HTML into Excel with Aspose.Cells C# | Aspose.Cells import HTML table and preserve line breaks in cell values | check if cell A1 contains newline after HtmlLoadOptions HTML import in C# | C# code to detect \n characters in a worksheet cell after HTML import using Aspose.Cells | convert HTML <br> elements to line feeds in Excel cells with Aspose.Cells .NET
// Tags: HtmlLoadOptions line break handling | Aspose.Cells HTML to Excel newline conversion | cell stringvalue newline verification C# | preserve br tags in worksheet cells | detect \n in imported HTML cell

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an HTML string containing <br> tags into an Aspose.Cells Workbook via HtmlLoadOptions, accesses cell A1, reads its StringValue, checks for '\n' newline characters, and prints the cell content along with a verification flag indicating whether line breaks were preserved.
class Program
{
    static void Main()
    {
        // HTML containing <br> tags
        string html = "<html><body><table><tr><td>Line1<br>Line2<br/>Line3</td></tr></table></body></html>";

        // Load the HTML into a workbook using HtmlLoadOptions
        var loadOptions = new HtmlLoadOptions();
        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            Workbook workbook = new Workbook(stream, loadOptions);

            // Get the first worksheet and the cell that contains the imported text
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"]; // Aspose.Cells places the first cell at A1

            // Retrieve the cell's string value
            string cellText = cell.StringValue;

            // Check if the cell text contains newline characters (line breaks)
            bool containsNewLine = cellText.Contains("\n");

            // Output the result
            Console.WriteLine("Cell content:");
            Console.WriteLine(cellText);
            Console.WriteLine("Contains newline characters: " + containsNewLine);
        }
    }
}
