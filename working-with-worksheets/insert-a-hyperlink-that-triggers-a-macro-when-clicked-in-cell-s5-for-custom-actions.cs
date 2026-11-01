// Title: Add a macro:// hyperlink to cell S5 and save the workbook as a macro‑enabled .xlsm file with Aspose.Cells for .NET
// AI Prompts: Create a macro-enabled workbook, write 'Run Macro' in cell S5, attach a macro://MyMacro hyperlink to that cell, and save as Xlsm using Aspose.Cells. | Modify the example to place a macro:// hyperlink in a different cell and reference a different macro name. | Generate C# code that adds multiple macro:// hyperlinks to various cells and saves the result as a macro‑enabled workbook.
// Common Searches: Aspose.Cells C# add macro:// hyperlink to a specific cell | how to programmatically create an .xlsm file with a macro link using Aspose.Cells | insert clickable macro hyperlink in Excel cell with Aspose.Cells .NET | save workbook as macro‑enabled Xlsm format after adding hyperlink in Aspose.Cells | C# example for linking a cell to an Excel VBA macro with Aspose.Cells
// Tags: macro hyperlink insertion Aspose.Cells | save workbook as xlsm Aspose.Cells | cell S5 hyperlink C# | macro:// scheme usage Excel | programmatic macro link creation Aspose.Cells

using System;
using Aspose.Cells;

// The code creates a new workbook, writes "Run Macro" into cell S5, adds a macro://MyMacro hyperlink to that cell, and saves the file as a macro‑enabled .xlsm workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put display text in cell S5
            Cell cell = sheet.Cells["S5"];
            cell.PutValue("Run Macro");

            // Add a hyperlink that triggers the macro named "MyMacro"
            // The "macro://" scheme tells Excel to execute the macro when the link is clicked.
            int row = cell.Row;
            int column = cell.Column;
            sheet.Hyperlinks.Add(row, column, row, column, "macro://MyMacro");

            // Save the workbook as a macro‑enabled file
            workbook.Save("HyperlinkMacro.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
