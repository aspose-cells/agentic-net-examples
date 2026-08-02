// Title: C# – Convert HTML with Formulas to Excel and Enable Automatic Calculation using Aspose.Cells
// Description: Shows how to load an HTML file that contains formulas with Aspose.Cells HtmlLoadOptions (LoadFormulas = true), set the workbook's calculation mode to Automatic, and save the output as an .xlsx file so formulas recalculate on open.
// Keywords: Aspose.Cells | HTML to Excel conversion | C# HtmlLoadOptions | LoadFormulas | automatic calculation mode | CalcModeType.Automatic | formula evaluation | convert HTML tables to XLSX | Aspose.Cells .NET | Excel workbook settings
// Common Searches: Aspose.Cells load HTML with formulas C# | set calculation mode automatic Aspose.Cells .NET | convert HTML table to Excel preserving formulas | C# example HTML to XLSX Aspose.Cells | enable automatic formula calculation after HTML import
// Developer Intent: Load an HTML document containing formulas, configure the workbook to recalculate automatically, and export it as an Excel workbook.
// Use Cases: Transform web‑generated reports that embed HTML tables with formulas into Excel files that update calculations instantly. | Batch‑process HTML invoices or statements that include embedded calculations, producing ready‑to‑analyze XLSX files. | Build a data‑ingestion pipeline that reads multiple HTML sources, sets automatic formula evaluation, and delivers Excel workbooks for downstream analytics.
// AI Prompts: Generate C# code using Aspose.Cells to load an HTML file with formulas, set the workbook's calculation mode to Automatic, and save it as an .xlsx file. | Explain how HtmlLoadOptions.LoadFormulas and Workbook.Settings.FormulaSettings.CalculationMode work together when converting HTML to Excel with Aspose.Cells. | Provide a step‑by‑step guide for converting a folder of HTML files to Excel workbooks that automatically evaluate formulas, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToExcel
{
    // Shows how to load an HTML file that contains formulas with Aspose.Cells HtmlLoadOptions (LoadFormulas = true), set the workbook's calculation mode to Automatic, and save the output as an .xlsx file so formulas recalculate on open.
    class Program
    {
        static void Main()
        {
            // Load the HTML file and import any formulas it contains
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                LoadFormulas = true // ensure formulas are recognized during load
            };
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Set the workbook calculation mode to Automatic so formulas are evaluated when needed
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Save the workbook as an Excel file
            workbook.Save("output.xlsx");
        }
    }
}
