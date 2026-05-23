using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a template workbook ----------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Header row
            cells["A1"].PutValue("Value1");
            cells["B1"].PutValue("Weight1");
            cells["C1"].PutValue("Value2");
            cells["D1"].PutValue("Weight2");
            cells["E1"].PutValue("WeightedAvg");

            // Template row with smart markers (will be expanded by WorkbookDesigner)
            cells["A2"].PutValue("&=Data.Value1");
            cells["B2"].PutValue("&=Data.Weight1");
            cells["C2"].PutValue("&=Data.Value2");
            cells["D2"].PutValue("&=Data.Weight2");
            cells["E2"].PutValue("&=Data.WeightedAvg"); // placeholder for the result

            // ---------- Prepare data source ----------
            List<RowData> data = new List<RowData>
            {
                new RowData { Value1 = 10, Weight1 = 2, Value2 = 20, Weight2 = 3 },
                new RowData { Value1 = 5,  Weight1 = 1, Value2 = 15, Weight2 = 4 },
                new RowData { Value1 = 8,  Weight1 = 2, Value2 = 12, Weight2 = 2 }
            };

            // ---------- Process smart markers ----------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = wb;               // assign workbook to designer
            designer.SetDataSource("Data", data); // bind data source
            designer.Process();                   // populate smart markers

            // ---------- Calculate weighted average per row ----------
            Worksheet resultWs = designer.Workbook.Worksheets[0];
            Cells resultCells = resultWs.Cells;

            int startRow = 1; // zero‑based index of first data row (row 2 in Excel)
            int endRow = startRow + data.Count - 1;

            for (int r = startRow; r <= endRow; r++)
            {
                // Formula: =(A*B + C*D) / (B + D) for the current row
                string formula = $"=({resultCells[r, 0].Name}*{resultCells[r, 1].Name}+{resultCells[r, 2].Name}*{resultCells[r, 3].Name})/({resultCells[r, 1].Name}+{resultCells[r, 3].Name})";

                // Set the formula in column E (index 4)
                resultCells[r, 4].SetFormula(formula, new FormulaParseOptions());
            }

            // ---------- Evaluate all formulas ----------
            designer.Workbook.CalculateFormula();

            // ---------- Save the result ----------
            string outputPath = "WeightedAverageResult.xlsx";
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Data class matching the smart marker names
    public class RowData
    {
        public double Value1 { get; set; }
        public double Weight1 { get; set; }
        public double Value2 { get; set; }
        public double Weight2 { get; set; }
        // Placeholder for the smart marker; its value will be overwritten by the formula
        public double WeightedAvg { get; set; }
    }
}