using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class MergeAndCalculate
    {
        static void Main()
        {
            try
            {
                const string destPath = "Destination.xlsx";
                const string srcPath = "Source.xlsx";
                const string resultPath = "MergedResult.xlsx";

                // Ensure input files exist
                if (!File.Exists(destPath))
                {
                    Console.WriteLine($"File not found: {destPath}");
                    return;
                }
                if (!File.Exists(srcPath))
                {
                    Console.WriteLine($"File not found: {srcPath}");
                    return;
                }

                // Load workbooks
                Workbook destWorkbook = new Workbook(destPath);
                Workbook srcWorkbook = new Workbook(srcPath);

                // Copy each worksheet from source to destination
                foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
                {
                    // AddCopy expects the sheet name, not the Worksheet object
                    destWorkbook.Worksheets.AddCopy(srcSheet.Name);
                }

                // Recalculate all formulas in the merged workbook
                destWorkbook.CalculateFormula();

                // Save the merged workbook
                destWorkbook.Save(resultPath);
                Console.WriteLine($"Merged workbook saved to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}