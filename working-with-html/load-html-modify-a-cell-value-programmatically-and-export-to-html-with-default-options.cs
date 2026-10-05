// Title: Load an HTML spreadsheet with Aspose.Cells, change cell A1, and save back to HTML using default options in C#
// AI Prompts: Read an HTML file into an Aspose.Cells Workbook, set cell A1 to a custom string, and write the workbook out as HTML with default settings using C#. | Using Aspose.Cells for .NET, load an existing HTML worksheet, update a cell value programmatically, and export the modified workbook back to HTML without custom save options. | Show how to modify a cell after loading an HTML spreadsheet with Aspose.Cells and then save the workbook as HTML in C#.
// Common Searches: asp.net aspose.cells load html file modify cell a1 and save as html | c# aspose.cells change value in html workbook then export to html | how to edit a cell in an html spreadsheet using Aspose.Cells C#
// Tags: load html workbook with Aspose.Cells | set cell value programmatically Aspose.Cells | export workbook to html default options | modify cell A1 in loaded html spreadsheet | c# aspose.cells html to html conversion

using System;
using Aspose.Cells;

// // This program loads an HTML file into an Aspose.Cells Workbook, changes cell A1 to "Hello World", and saves the workbook back to HTML using the default SaveFormat.Html options.
class Program
{
    static void Main()
    {
        // Input HTML file path
        string inputHtml = "input.html";
        // Output HTML file path
        string outputHtml = "output.html";

        // Load the HTML document into a Workbook
        Workbook workbook = new Workbook(inputHtml);

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify a cell value (e.g., set A1 to "Hello World")
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Hello World");

        // Export the workbook back to HTML using default options
        workbook.Save(outputHtml, SaveFormat.Html);
    }
}
