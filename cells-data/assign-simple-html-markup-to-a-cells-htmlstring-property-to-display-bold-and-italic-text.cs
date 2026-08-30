// Title: Assign bold and italic HTML markup to a cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, sets the HtmlString of cell A1 to include <b> and <i> tags, and saves the file as XLSX with Aspose.Cells. | Demonstrate how to use the Aspose.Cells HtmlString property to render bold and italic HTML text inside an Excel worksheet cell in a .NET application.
// Common Searches: Aspose.Cells C# how to display bold and italic text in an Excel cell using HtmlString | example of setting HtmlString property for a worksheet cell in .NET | render HTML tags inside Excel cells with Aspose.Cells library | save Excel file with HTML-formatted cell content using Aspose.Cells C#
// Tags: Aspose.Cells HtmlString usage | C# set HTML content in Excel cell | format cell bold italic with Aspose | save workbook as XLSX with HTML markup | Excel cell HTML rendering .NET

using System;
using Aspose.Cells;

// Creates a workbook, assigns bold and italic HTML markup to cell A1 via the HtmlString property, and saves the result as BoldItalicDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Assign HTML markup to cell A1 to display bold and italic text
        worksheet.Cells["A1"].HtmlString = "This is <b>bold</b> and <i>italic</i> text";

        // Save the workbook to an XLSX file
        workbook.Save("BoldItalicDemo.xlsx", SaveFormat.Xlsx);
    }
}
