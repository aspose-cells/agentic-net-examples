// Title: Calculate Weighted Average with Smart Markers and SUMPRODUCT in Aspose.Cells for .NET
// Description: This C# example shows how to create an Excel workbook, embed smart markers for Item, Value, and Weight, bind a DataTable, process the markers, insert a dynamic SUMPRODUCT/SUM formula that adapts to the populated rows, evaluate the formula with Worksheet.CalculateFormula, display the result, and save the file as WeightedAverageResult.xlsx.
// Keywords: Aspose.Cells | C# | smart markers | weighted average | SUMPRODUCT | Worksheet.CalculateFormula | WorkbookDesigner | dynamic Excel formula | DataTable binding | Excel automation
// Common Searches: Aspose.Cells weighted average smart markers | C# SUMPRODUCT formula after processing smart markers | calculate Excel formula programmatically with Aspose.Cells | how to use Worksheet.CalculateFormula in .NET | dynamic row range in Aspose.Cells formula
// Developer Intent: Generate an Excel report where rows are filled via smart markers and a weighted average is computed automatically with a formula that references the inserted data.
// Use Cases: Create a reusable template that expands rows from a DataTable and returns the weighted average without manual editing. | Build a financial or inventory report where each item’s value and weight are supplied by code and the summary cell updates instantly. | Integrate the calculated weighted average into further C# logic by retrieving the value from Worksheet.CalculateFormula.
// AI Prompts: Modify the formula to exclude rows with zero or missing weight. | Add code to format the result cell as a number with two decimal places. | Show how to calculate separate weighted averages for multiple categories using additional smart markers.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsWeightedAverageDemo
{
    // This C# example shows how to create an Excel workbook, embed smart markers for Item, Value, and Weight, bind a DataTable, process the markers, insert a dynamic SUMPRODUCT/SUM formula that adapts to the populated rows, evaluate the formula with Worksheet.CalculateFormula, display the result, and save the file as WeightedAverageResult.xlsx.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (lifecycle rule: create)
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 2. Build a template with smart markers
            // -------------------------------------------------
            // Header row
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("Weight");

            // Data rows using smart markers (the marker name is "Data")
            // &="Data.Value" and &="Data.Weight" will be replaced by actual values
            cells["A2"].PutValue("&=Data.Item");
            cells["B2"].PutValue("&=Data.Value");
            cells["C2"].PutValue("&=Data.Weight");

            // -------------------------------------------------
            // 3. Prepare the data source (DataTable)
            // -------------------------------------------------
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Item", typeof(string));
            dt.Columns.Add("Value", typeof(double));
            dt.Columns.Add("Weight", typeof(double));

            dt.Rows.Add("A", 10.0, 2.0);
            dt.Rows.Add("B", 20.0, 3.0);
            dt.Rows.Add("C", 30.0, 5.0);
            // Add more rows as needed...

            // -------------------------------------------------
            // 4. Bind data source and process smart markers
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // rule: Process()

            // -------------------------------------------------
            // 5. Insert a formula that calculates the weighted average
            // -------------------------------------------------
            // Determine the last data row after processing
            int lastRow = cells.MaxDataRow; // includes header, so data starts at row 2

            // Build the formula string using SUMPRODUCT and SUM
            // Example: =SUMPRODUCT(B2:B4, C2:C4) / SUM(C2:C4)
            string formula = $"=SUMPRODUCT(B2:B{lastRow},C2:C{lastRow})/SUM(C2:C{lastRow})";

            // Place the formula in D2 (or any cell you prefer)
            cells[$"D2"].Formula = formula;

            // -------------------------------------------------
            // 6. Calculate the formula using Worksheet.CalculateFormula(string)
            // -------------------------------------------------
            // The CalculateFormula method returns the computed value.
            object weightedAvg = sheet.CalculateFormula(formula);

            // Also calculate all formulas in the workbook to update cell values
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 7. Output the result (for demonstration)
            // -------------------------------------------------
            Console.WriteLine($"Weighted Average (calculated via Worksheet.CalculateFormula): {weightedAvg}");
            Console.WriteLine($"Weighted Average (value stored in D2 after workbook.CalculateFormula): {cells[1, 3].Value}");

            // -------------------------------------------------
            // 8. Save the workbook (lifecycle rule: save)
            // -------------------------------------------------
            workbook.Save("WeightedAverageResult.xlsx");
        }
    }
}
