using System;
using Aspose.Cells;

namespace HtmlToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Load the HTML file with options to import formulas
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                LoadFormulas = true   // Ensure formulas present in HTML are loaded as formulas
            };
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Set the workbook calculation mode to Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook as an Excel file
            string excelPath = "output.xlsx";
            workbook.Save(excelPath);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel '{excelPath}' with automatic calculation mode.");
        }
    }
}