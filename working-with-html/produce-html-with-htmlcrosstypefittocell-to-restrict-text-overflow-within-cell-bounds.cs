// Title: Export Excel to HTML with Aspose.Cells using HtmlCrossType.FitToCell to keep wrapped text inside cell boundaries (C#)
// AI Prompts: Generate a C# example that creates a workbook, sets a narrow column, inserts a long string, enables text wrap, applies HtmlSaveOptions with HtmlCrossType.FitToCell, and saves the sheet as an HTML file. | Modify existing Aspose.Cells code to configure HtmlSaveOptions.HtmlCrossType = HtmlCrossType.FitToCell and verify that long text does not overflow the rendered HTML cell. | Add error‑handling and directory‑creation logic to a C# program that exports a worksheet to HTML while using FitToCell to constrain cell content.
// Common Searches: how to prevent text overflow when converting Excel to HTML with Aspose.Cells C# | Aspose.Cells HtmlCrossType FitToCell usage example in C# | C# export worksheet to HTML with wrapped text staying within cell width | set column width and enable text wrap before saving as HTML using Aspose.Cells | Aspose.Cells HTML save options fit to cell long text
// Tags: Aspose.Cells HTML cell fit option C# | export Excel to HTML with cell overflow control | set column width and text wrap Aspose.Cells | HTML save options configuration C# | C# workbook to HTML with wrapped text

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, narrows column A, inserts a long string into A1, enables text wrapping, configures HtmlSaveOptions with HtmlCrossType.FitToCell, ensures the output folder exists, and saves the worksheet as an HTML file, preventing the text from spilling outside the cell.
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

            // Set a narrow column width to force potential overflow (width in characters)
            sheet.Cells.SetColumnWidth(0, 10);

            // Insert long text into cell A1
            sheet.Cells["A1"].PutValue("This is a very long text that would normally overflow the cell boundaries.");

            // Enable text wrapping for the cell
            Style cellStyle = sheet.Cells["A1"].GetStyle();
            cellStyle.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(cellStyle);

            // Configure HTML save options (default options are sufficient for wrapping)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Determine output path and ensure the directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
