// Title: C# – Convert HTML (including nested tables) to Excel with Aspose.Cells – Separate worksheets per table
// Description: Loads an HTML file that may contain nested <table> elements using Aspose.Cells HtmlLoadOptions, imports each table as an individual worksheet, renames the sheets to reflect their hierarchy (Table_Level_1, Table_Level_2, …), and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells HTML to Excel C# | nested HTML tables conversion | HtmlLoadOptions separate worksheets | TableLoadOptions mapping | rename worksheets by level | convert input.html to output.xlsx
// Common Searches: Aspose.Cells load each HTML table into its own worksheet | C# convert nested HTML tables to separate Excel sheets | how to rename Excel worksheets based on HTML table hierarchy | map HTML table index to worksheet index Aspose.Cells
// Developer Intent: Generate an Excel workbook where every HTML table, including nested ones, is placed on a distinct worksheet named according to its hierarchy level.
// Use Cases: Transform a complex HTML report with multiple nested tables into an Excel file where each table resides on a separate sheet for easier data analysis. | Export HTML email templates to Excel while preserving the original table structure across individual worksheets. | Create an automated data‑extraction pipeline that reads web pages and outputs level‑based worksheets for downstream processing.
// AI Prompts: Write C# code with Aspose.Cells that converts an HTML document containing nested tables into an XLSX workbook, placing each table on a separate worksheet named Table_Level_1, Table_Level_2, etc. | Explain how HtmlLoadOptions.TableLoadOptions can map specific HTML table indexes to particular worksheet positions when loading HTML into a Workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an HTML file that may contain nested <table> elements using Aspose.Cells HtmlLoadOptions, imports each table as an individual worksheet, renames the sheets to reflect their hierarchy (Table_Level_1, Table_Level_2, …), and saves the workbook as an XLSX file.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Input HTML file that may contain nested tables
        string htmlPath = "input.html";

        // Desired output Excel file
        string excelPath = "output.xlsx";

        // Load options for HTML – each <table> element will be imported as a separate worksheet.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // (Optional) Explicitly map table indexes to worksheet indexes.
        // For example, map the first three tables to worksheets 0, 1, and 2.
        // loadOptions.TableLoadOptions.Add(0, 0);
        // loadOptions.TableLoadOptions.Add(1, 1);
        // loadOptions.TableLoadOptions.Add(2, 2);

        // Load the HTML document into a Workbook using the load options.
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Rename worksheets to indicate their hierarchical level (Level 1, Level 2, ...).
        // This helps identify which worksheet originated from which table.
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.Name = $"Table_Level_{i + 1}";
        }

        // Save the workbook as an Excel file.
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}
