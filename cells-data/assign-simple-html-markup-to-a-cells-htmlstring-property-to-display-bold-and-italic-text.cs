// Title: Set bold and italic HTML markup in an Excel cell with Aspose.Cells for .NET
// Description: Demonstrates how to assign an HTML string containing <b> and <i> tags to a worksheet cell using the HtmlString property, then save the workbook as XLSX.
// Keywords: Aspose.Cells | .NET | HtmlString | HTML formatting | bold text | italic text | C# example | Excel cell styling
// Common Searches: Aspose.Cells HtmlString bold italic | C# set HTML in Excel cell Aspose | How to use HtmlString property in Aspose.Cells | Apply HTML tags to Excel cell with Aspose.Cells | Display formatted text in Excel using Aspose.Cells
// Developer Intent: Apply HTML markup to a cell so the displayed value shows bold and italic formatting.
// Use Cases: Create report headers with mixed styling without manual cell formatting. | Export data where certain keywords need emphasis directly from HTML sources. | Build templates that render HTML‑styled content in Excel worksheets.
// AI Prompts: Show a C# example that adds underline and color HTML tags to a cell using Aspose.Cells HtmlString. | Explain how to load an HTML string from a database and assign it to a worksheet cell with HtmlString.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlStringDemo
{
    // Demonstrates how to assign an HTML string containing <b> and <i> tags to a worksheet cell using the HtmlString property, then save the workbook as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign HTML markup to cell A1.
            // The markup includes bold (<b>) and italic (<i>) tags.
            worksheet.Cells["A1"].HtmlString = "This is <b>bold</b> and <i>italic</i> text";

            // Save the workbook to an XLSX file
            workbook.Save("HtmlStringDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
