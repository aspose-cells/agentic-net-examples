using System;
using Aspose.Cells;

namespace AsposeCellsFindDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add sample data and formulas
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Plain values
            sheet.Cells["A1"].PutValue("Alpha");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["C1"].PutValue(DateTime.Today);

            // Formulas referencing other cells
            sheet.Cells["A2"].Formula = "=A1 & \"_suffix\"";
            sheet.Cells["B2"].Formula = "=B1 * 2";
            sheet.Cells["C2"].Formula = "=TODAY()";

            // Save the workbook in XLSX format
            string filePath = "FindDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 2. Load the workbook (parsing formulas on open)
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = true // ensure formulas are parsed when loading
            };
            Workbook loadedWb = new Workbook(filePath, loadOptions);
            Worksheet loadedSheet = loadedWb.Worksheets[0];

            // -----------------------------------------------------------------
            // 3. Find a cell that contains a specific text inside a formula
            //    (e.g., searching for the reference "A1" within formulas)
            // -----------------------------------------------------------------
            FindOptions formulaSearchOptions = new FindOptions
            {
                LookInType = LookInType.OnlyFormulas, // search only in formulas
                LookAtType = LookAtType.Contains,     // partial match
                CaseSensitive = false
            };

            Cell formulaCell = loadedSheet.Cells.Find("A1", null, formulaSearchOptions);
            if (formulaCell != null)
            {
                Console.WriteLine($"Formula match found at {formulaCell.Name}");
                Console.WriteLine($"Formula: {formulaCell.Formula}");
            }
            else
            {
                Console.WriteLine("No formula containing 'A1' was found.");
            }

            // -----------------------------------------------------------------
            // 4. Find a cell that contains a specific value (e.g., the number 123)
            // -----------------------------------------------------------------
            FindOptions valueSearchOptions = new FindOptions
            {
                LookInType = LookInType.Values,      // search in cell values
                LookAtType = LookAtType.EntireContent, // exact match
                CaseSensitive = false
            };

            Cell valueCell = loadedSheet.Cells.Find(123, null, valueSearchOptions);
            if (valueCell != null)
            {
                Console.WriteLine($"Value match found at {valueCell.Name}");
                Console.WriteLine($"Cell value: {valueCell.Value}");
            }
            else
            {
                Console.WriteLine("No cell with value 123 was found.");
            }

            // -----------------------------------------------------------------
            // 5. Demonstrate searching within a specific range
            // -----------------------------------------------------------------
            CellArea searchArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 2,
                EndColumn = 2
            };
            FindOptions rangeSearchOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                SearchOrderByRows = true
            };
            rangeSearchOptions.SetRange(searchArea);

            Cell rangeCell = loadedSheet.Cells.Find("Alpha", null, rangeSearchOptions);
            if (rangeCell != null)
            {
                Console.WriteLine($"Range search found at {rangeCell.Name}");
            }
            else
            {
                Console.WriteLine("Range search did not find the target.");
            }
        }
    }
}