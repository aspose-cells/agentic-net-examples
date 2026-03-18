using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourceFile = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourceFile);

        // Access the first worksheet (or any specific worksheet as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the defined print area for the worksheet
        string printArea = worksheet.PageSetup.PrintArea;
        Console.WriteLine($"Defined Print Area: {printArea}");

        // Configure HTML save options to export only the print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportPrintAreaOnly = true;   // Use the rule to export only the print area
        htmlOptions.ExportGridLines = true;       // Optional: include grid lines for better layout

        // Save the workbook as an HTML file preserving the layout of the print area
        string htmlFile = "output.html";
        workbook.Save(htmlFile, htmlOptions);

        Console.WriteLine($"HTML file generated: {htmlFile}");
    }
}