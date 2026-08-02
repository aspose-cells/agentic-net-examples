// Title: C# – Convert HTML Table colspan/rowspan to Merged Cells in Excel with Aspose.Cells
// Description: Demonstrates loading an HTML file via Aspose.Cells HtmlLoadOptions, automatically turning <td colspan> and <td rowspan> elements into merged cells in the workbook, optionally selecting tables by index, and saving the result as an XLSX file.
// Keywords: Aspose.Cells | C# | HTML to Excel | colspan | rowspan | merged cells | HtmlLoadOptions | table import | Excel export
// Common Searches: Aspose.Cells convert HTML table colspan to merged cells | C# load HTML with rowspan into Excel | preserve merged cells when importing HTML to Excel using Aspose.Cells | select specific HTML tables for conversion Aspose.Cells | HtmlLoadOptions TableLoadOptions example
// Developer Intent: Import an HTML document and have any colspan or rowspan cells become merged cells in the generated Excel workbook.
// Use Cases: Migrate web‑based reports with complex tables to Excel while retaining the original merged‑cell layout. | Extract selected tables from an HTML page into separate worksheets, preserving colspan/rowspan as merged cells. | Automate conversion of HTML email newsletters containing tables into Excel spreadsheets, ensuring merged cells match the source markup.
// AI Prompts: Write C# code that loads an HTML file with multiple tables, converts each table into its own worksheet, and keeps colspan/rowspan as merged cells using Aspose.Cells. | Explain how HtmlLoadOptions.TableLoadOptions can be used to import only tables with specific IDs or indexes from an HTML document. | Show how to apply a custom style (e.g., background color) to all merged cells after loading the HTML with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates loading an HTML file via Aspose.Cells HtmlLoadOptions, automatically turning <td colspan> and <td rowspan> elements into merged cells in the workbook, optionally selecting tables by index, and saving the result as an XLSX file.
class HtmlTableToMergedCells
{
    static void Main()
    {
        // Load the HTML file. Aspose.Cells automatically converts <td colspan> and <td rowspan>
        // into merged cells in the resulting worksheet.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // If you need to import specific tables only, you can add their indexes or IDs:
        // loadOptions.TableLoadOptions.Add(0); // import the first table
        // loadOptions.TableLoadOptions.Add(1); // import the second table, etc.

        // Load the HTML document into a workbook.
        Workbook workbook = new Workbook("input.html", loadOptions);

        // The worksheet now contains merged cells that correspond to the original colspan/rowspan.
        // Save the workbook to an Excel file.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
