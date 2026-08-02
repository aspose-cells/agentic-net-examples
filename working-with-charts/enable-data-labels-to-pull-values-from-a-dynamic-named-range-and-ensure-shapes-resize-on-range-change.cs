using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDynamicDataLabels
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for categories and values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Populate label values that will be used for data labels
            // This column will be expanded later to demonstrate dynamic range handling
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Ten");
            sheet.Cells["C3"].PutValue("Twenty");
            sheet.Cells["C4"].PutValue("Thirty");

            // Define a dynamic named range "LabelValues" that refers to column C starting at C2
            // The range expands automatically based on the number of non‑empty cells in column C
            int nameIndex = workbook.Worksheets.Names.Add("LabelValues");
            Name labelName = workbook.Worksheets.Names[nameIndex];
            // OFFSET formula: start at C2, height = COUNTA(C:C)-1 (exclude header), width = 1
            labelName.RefersTo = "=Sheet1!$C$2:OFFSET(Sheet1!$C$2,COUNTA(Sheet1!$C:$C)-2,0)";

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set series data and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to pull values from the dynamic named range
            Series series = chart.NSeries[0];
            series.DataLabels.ShowCellRange = true;          // Show cell range as label text
            series.DataLabels.LinkedSource = "LabelValues"; // Link to the dynamic named range
            series.DataLabels.IsResizeShapeToFitText = true; // Auto‑fit shape to text

            // Initial calculation so that the chart reflects the current data
            chart.Calculate();

            // ----- Simulate a change that expands the dynamic range -----
            // Add a new category and value
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);
            // Add a corresponding label
            sheet.Cells["C5"].PutValue("Forty");

            // Refresh dynamic array formulas (if any) and recalculate the workbook
            workbook.RefreshDynamicArrayFormulas(true);
            workbook.CalculateFormula();

            // Re‑calculate the chart so that data labels pick up the new label value
            chart.Calculate();

            // Save the workbook
            workbook.Save("DynamicDataLabels.xlsx");
        }
    }
}