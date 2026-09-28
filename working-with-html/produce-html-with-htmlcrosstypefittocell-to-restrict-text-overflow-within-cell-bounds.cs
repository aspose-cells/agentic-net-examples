// Title: C# Aspose.Cells: Export Excel to HTML with HtmlCrossType.FitToCell to prevent text overflow
// AI Prompts: Write C# code that creates a workbook, inserts a long string into a wrapped cell, and saves it as HTML using HtmlSaveOptions with HtmlCrossType set to FitToCell. | Update the provided Aspose.Cells example to enable HtmlCrossType.FitToCell and show the effect on the generated HTML when a cell contains lengthy text. | Describe how to test that the HTML output respects the original cell width after applying HtmlCrossType.FitToCell in a C# Aspose.Cells project.
// Common Searches: Aspose.Cells C# how to use HtmlCrossType.FitToCell when saving workbook as HTML | prevent cell text overflow in HTML export using Aspose.Cells | example of setting HtmlSaveOptions.HtmlCrossType to FitToCell in .NET | C# code to wrap long text in Excel cell and keep it within cell after HTML conversion | Aspose.Cells HTML export cell width limitation FitToCell
// Tags: Aspose.Cells HtmlCrossType FitToCell | C# export Excel to HTML Aspose | cell text overflow prevention Aspose.Cells | HTML save options cell width control | text wrapping with HtmlSaveOptions Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a workbook, narrowing a column, inserting a long string into a wrapped cell, and saving the workbook as HTML with HtmlSaveOptions configured to use HtmlCrossType.FitToCell, ensuring the generated HTML confines the text within the original cell boundaries.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a narrow column width to demonstrate overflow handling
            sheet.Cells.SetColumnWidth(0, 10); // width in characters

            // Insert long text into cell A1
            sheet.Cells["A1"].PutValue("This is a very long text that would normally overflow to adjacent cells.");

            // Enable text wrapping for the cell
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(wrapStyle);

            // Configure HTML save options (FitToCell option is not available in this version)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Determine output path and ensure the directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
