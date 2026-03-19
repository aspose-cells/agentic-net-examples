using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbookEnglishFormulasDemo
    {
        public static void Run()
        {
            // Create load options and set the culture to English (en-US)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("en-US"); // forces English formula parsing

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Access a cell that contains a formula
            Cell cell = workbook.Worksheets[0].Cells["A1"];

            // Output the formula in standard (English) format
            Console.WriteLine("Standard (English) Formula: " + cell.Formula);

            // Output the localized formula (if any) for comparison
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookEnglishFormulasDemo.Run();
        }
    }
}