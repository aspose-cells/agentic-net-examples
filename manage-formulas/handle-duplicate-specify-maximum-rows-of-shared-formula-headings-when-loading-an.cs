using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HandleDuplicateSharedFormulaHeadings
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);
            workbook.Settings.MaxRowsOfSharedFormula = 1024;
            workbook.CalculateFormula();
            workbook.Save(outputPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HandleDuplicateSharedFormulaHeadings.Run();
        }
    }
}