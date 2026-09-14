// Title: Load an HTML spreadsheet, change a cell value, and save it back to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Read an HTML file into a Workbook, set cell A1 to a new string, and write the workbook out as HTML with default settings in C#. | Generate C# code that opens an HTML spreadsheet, updates a chosen cell, and saves the result back to HTML using Aspose.Cells. | Show how to programmatically edit a cell in an HTML workbook and export it without customizing the HtmlSaveOptions.
// Common Searches: Aspose.Cells C# load html file edit cell and save as html example | how to change a cell value in an html spreadsheet using Aspose.Cells for .NET | C# code to modify A1 in an html workbook and export with default options | default HtmlSaveOptions when saving a workbook to HTML with Aspose.Cells | update html spreadsheet cell programmatically Aspose.Cells tutorial
// Tags: html workbook loading Aspose.Cells | cell value update C# Aspose.Cells | default html export Aspose.Cells | html spreadsheet modification Aspose.Cells | Aspose.Cells html-to-html conversion

using System;
using Aspose.Cells;

// The program loads an HTML file into an Aspose.Cells Workbook, changes the value of cell A1, and saves the workbook back to HTML using the default SaveFormat.Html settings.
class Program
{
    static void Main()
    {
        // Load the source HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Modify the value of cell A1 (you can change the address as needed)
        sheet.Cells["A1"].PutValue("Modified Value");

        // Save the workbook back to HTML using default export options
        workbook.Save("output.html", SaveFormat.Html);
    }
}
