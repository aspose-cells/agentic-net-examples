using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file that already contains formulas.
            string sourcePath = "TemplateWithFormulas.xlsx";

            // -----------------------------------------------------------------
            // 1. Load the workbook with specific LoadOptions.
            //    ParsingFormulaOnOpen = false skips parsing during load,
            //    which can improve performance for large files.
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // -----------------------------------------------------------------
            // 2. (Optional) Use WorkbookDesigner to bind a data source.
            //    Setting CalculateFormula = true makes the designer evaluate
            //    formulas after data binding.
            // -----------------------------------------------------------------
            DataTable dt = new DataTable("Sales");
            dt.Columns.Add("Product");
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add("Apple", 10, 1.5);
            dt.Rows.Add("Banana", 20, 0.9);
            dt.Rows.Add("Cherry", 15, 2.2);

            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.CalculateFormula = true; // evaluate formulas after binding
            designer.Process(); // apply the data source to the template

            // -----------------------------------------------------------------
            // 3. Parse formulas that were not parsed during the load step.
            //    This is required because we set ParsingFormulaOnOpen = false.
            // -----------------------------------------------------------------
            workbook.ParseFormulas(false);

            // -----------------------------------------------------------------
            // 4. Calculate all formulas in the workbook.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 5. (Optional) Retrieve a formula from the workbook using AI.
            //    The Aspose.Cells.AI namespace is not available in the standard
            //    Aspose.Cells package, so this step is omitted or can be
            //    implemented using a custom solution.
            // -----------------------------------------------------------------
            // string question = "What is the total sales amount?";
            // string formulaFromAI = CellsAI.GetExcelFormula(sourcePath, question);
            // Console.WriteLine($"AI‑generated formula for \"{question}\": {formulaFromAI}");

            // -----------------------------------------------------------------
            // 6. Save the processed workbook.
            // -----------------------------------------------------------------
            string outputPath = "ProcessedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
    }
}