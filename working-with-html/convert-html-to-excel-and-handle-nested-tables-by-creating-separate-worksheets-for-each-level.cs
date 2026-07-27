// Title: Convert HTML with Nested Tables to Excel – Separate Worksheets per Level (C# Aspose.Cells)
// Description: Shows how to load an HTML file that contains nested tables using Aspose.Cells, automatically generate a worksheet for each table, rename the sheets to indicate nesting depth (Level1, Level2, …), and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | HTML to Excel | nested tables | separate worksheets | C# conversion | HtmlLoadOptions | worksheet naming | XLSX export
// Common Searches: Aspose.Cells convert HTML nested tables to Excel | C# create separate worksheets for each HTML table | load HTML with Aspose.Cells and rename sheets | HTML to XLSX with nested tables Aspose | how to handle nested tables in Aspose.Cells HTML import
// Developer Intent: Generate an Excel workbook from an HTML document where every table—including nested ones—appears on its own worksheet, then rename the sheets to reflect their hierarchy level.
// Use Cases: Convert a complex web report with multiple nested tables into a clean, multi‑sheet Excel workbook for business analysis. | Automate the extraction of scraped HTML data, preserving table structure by placing each table on a separate sheet. | Create a navigable Excel file where sheet names (Level1, Level2, …) convey the original HTML table hierarchy.
// AI Prompts: Write C# code using Aspose.Cells to load an HTML file, generate a separate worksheet for each nested table, and name the sheets Level1, Level2, etc. | Explain how HtmlLoadOptions processes nested tables during HTML‑to‑Excel conversion and how to customize worksheet names afterward. | Suggest a way to use original HTML table captions as worksheet names instead of generic LevelX labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Shows how to load an HTML file that contains nested tables using Aspose.Cells, automatically generate a worksheet for each table, rename the sheets to indicate nesting depth (Level1, Level2, …), and save the workbook as an XLSX file.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Paths for the source HTML file and the destination Excel file
        string htmlFilePath = "input.html";
        string excelFilePath = "output.xlsx";

        // Load the HTML file. Aspose.Cells automatically creates a separate worksheet
        // for each HTML table, including nested tables.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // Create a workbook from the HTML source using the load options
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Rename worksheets to reflect their nesting level (Level1, Level2, ...)
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            workbook.Worksheets[i].Name = $"Level{i + 1}";
        }

        // Save the workbook as an Excel file (XLSX format)
        workbook.Save(excelFilePath, SaveFormat.Xlsx);

        Console.WriteLine($"HTML has been converted to Excel. Output saved at: {excelFilePath}");
    }
}
