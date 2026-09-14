// Title: Calculate weighted average per row with smart markers and Worksheet.CalculateFormula in Aspose.Cells for C#
// AI Prompts: Generate C# code that adds smart markers for value and weight columns, binds a DataTable, processes the markers, builds a weighted‑average formula for each row, evaluates it with Worksheet.CalculateFormula, and writes the result to a new column. | Modify the example to compute weighted averages for three value/weight pairs per row, inserting the formula into the worksheet instead of using CalculateFormula directly. | Show how to export the workbook with the calculated weighted averages to an XLSX file and verify the results programmatically.
// Common Searches: how to use smart markers to calculate weighted average in Aspose.Cells C# | Worksheet.CalculateFormula with dynamic cell references example Aspose.Cells | binding a DataTable to smart markers for per‑row calculations Aspose.Cells | C# Aspose.Cells compute weighted average across multiple columns | Aspose.Cells calculate formula without placing it in a cell
// Tags: smart markers weighted average calculation | Worksheet.CalculateFormula dynamic expression | C# Aspose.Cells DataTable binding | per‑row formula generation Aspose.Cells | save workbook with calculated results Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsWeightedAverageDemo
{
    // The sample creates a workbook, defines headers, inserts smart markers for value and weight fields, binds a DataTable as the data source, processes the markers, then iterates over each data row to construct and evaluate a weighted‑average formula using Worksheet.CalculateFormula. The computed result is written to column E and the workbook is saved as WeightedAverageResult.xlsx.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 2. Define headers
            // ------------------------------------------------------------
            cells["A1"].PutValue("Value1");
            cells["B1"].PutValue("Weight1");
            cells["C1"].PutValue("Value2");
            cells["D1"].PutValue("Weight2");
            cells["E1"].PutValue("WeightedAvg");

            // ------------------------------------------------------------
            // 3. Insert smart markers – these will be replaced by data source values
            // ------------------------------------------------------------
            // Row 2
            cells["A2"].PutValue("&=Data.Value1");
            cells["B2"].PutValue("&=Data.Weight1");
            cells["C2"].PutValue("&=Data.Value2");
            cells["D2"].PutValue("&=Data.Weight2");
            // Row 3
            cells["A3"].PutValue("&=Data.Value1");
            cells["B3"].PutValue("&=Data.Weight1");
            cells["C3"].PutValue("&=Data.Value2");
            cells["D3"].PutValue("&=Data.Weight2");
            // Row 4
            cells["A4"].PutValue("&=Data.Value1");
            cells["B4"].PutValue("&=Data.Weight1");
            cells["C4"].PutValue("&=Data.Value2");
            cells["D4"].PutValue("&=Data.Weight2");

            // ------------------------------------------------------------
            // 4. Prepare a DataTable as the smart marker data source
            // ------------------------------------------------------------
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Value1", typeof(double));
            dt.Columns.Add("Weight1", typeof(double));
            dt.Columns.Add("Value2", typeof(double));
            dt.Columns.Add("Weight2", typeof(double));

            // Sample rows
            dt.Rows.Add(10, 2, 30, 3);   // Weighted avg = (10*2 + 30*3) / (2+3) = 22
            dt.Rows.Add(20, 4, 40, 1);   // Weighted avg = (20*4 + 40*1) / (4+1) = 24
            dt.Rows.Add(5,  5, 15, 5);   // Weighted avg = (5*5 + 15*5) / (5+5) = 10

            // ------------------------------------------------------------
            // 5. Bind the data source to the workbook designer and process smart markers
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();          // process

            // ------------------------------------------------------------
            // 6. Calculate weighted averages using Worksheet.CalculateFormula(string)
            //    The formula references the cells that were filled by smart markers.
            // ------------------------------------------------------------
            for (int row = 2; row <= 4; row++)
            {
                // Build the formula string for the current row
                // Example for row 2: =(A2*B2 + C2*D2) / (B2 + D2)
                string formula = $"=({cells[row, 0].Name}*{cells[row, 1].Name} + " +
                                 $"{cells[row, 2].Name}*{cells[row, 3].Name}) / " +
                                 $"({cells[row, 1].Name} + {cells[row, 3].Name})";

                // Calculate the formula directly (no need to place it in a cell)
                object result = sheet.CalculateFormula(formula);

                // Store the result in column E of the same row
                cells[row, 4].PutValue(result);
            }

            // ------------------------------------------------------------
            // 7. Save the workbook (output)
            // ------------------------------------------------------------
            workbook.Save("WeightedAverageResult.xlsx");   // save
        }
    }
}
