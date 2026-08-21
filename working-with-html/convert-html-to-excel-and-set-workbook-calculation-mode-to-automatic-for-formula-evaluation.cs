// Title: C# – Convert HTML to Excel with Automatic Formula Evaluation using Aspise.Cells
// Description: Loads an HTML file with Aspose.Cells, preserves embedded formulas, switches the workbook to automatic calculation mode, optionally forces an immediate recalculation, and saves the result as an XLSX file.
// Keywords: Aspose.Cells HTML to Excel | C# convert HTML to XLSX | load formulas from HTML | automatic formula calculation | CalcModeType.Automatic | Aspose.Cells workbook settings
// Common Searches: Aspose.Cells load HTML with formulas C# | set workbook calculation mode to automatic Aspose.Cells | convert HTML report to Excel preserving formulas | how to recalculate formulas after HTML import Aspose
// Developer Intent: Import an HTML document, keep its formulas intact, enable auto‑recalculation, and export to Excel.
// Use Cases: Transform web‑based reports that contain formulas into editable Excel workbooks. | Batch‑process multiple HTML files, preserving calculations for downstream analysis. | Create a data‑pipeline that converts HTML tables with embedded formulas into ready‑to‑use XLSX files.
// AI Prompts: Write C# code with Aspose.Cells that loads an HTML file, retains its formulas, sets automatic calculation, and saves as .xlsx. | Explain how to enable auto‑recalculation in Aspose.Cells after importing HTML and how to trigger an immediate formula evaluation. | Provide a step‑by‑step tutorial for converting HTML reports to Excel in C#, ensuring formulas are preserved and evaluated automatically.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlToExcel
{
    // Loads an HTML file with Aspose.Cells, preserves embedded formulas, switches the workbook to automatic calculation mode, optionally forces an immediate recalculation, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Path for the resulting Excel file
            string excelPath = "output.xlsx";

            // Load the HTML file with options that import formulas
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                // Ensure that any formulas present in the HTML are loaded as formulas
                LoadFormulas = true
            };

            // Create a workbook from the HTML source using the specified load options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Set the workbook calculation mode to Automatic so that formulas are evaluated
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Optionally calculate all formulas immediately (not required for the mode,
            // but ensures values are up‑to‑date before saving)
            workbook.CalculateFormula();

            // Save the workbook as an Excel file
            workbook.Save(excelPath);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel '{excelPath}' with automatic calculation mode.");
        }
    }
}
