using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicSecondaryAxis
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Primary categories (A column)
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("Cat1");
            cells["A3"].PutValue("Cat2");
            cells["A4"].PutValue("Cat3");
            cells["A5"].PutValue("Cat4");

            // Primary values (B column)
            cells["B1"].PutValue("Primary Values");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);
            cells["B5"].PutValue(40);

            // Secondary categories (C column)
            cells["C1"].PutValue("SecCategory");
            cells["C2"].PutValue("SubCat1");
            cells["C3"].PutValue("SubCat2");
            cells["C4"].PutValue("SubCat3");
            cells["C5"].PutValue("SubCat4");

            // Secondary values (D column)
            cells["D1"].PutValue("Secondary Values");
            cells["D2"].PutValue(100);
            cells["D3"].PutValue(200);
            cells["D4"].PutValue(300);
            cells["D5"].PutValue(400);

            // Lookup criteria for secondary category (G1)
            cells["G1"].PutValue("SubCat3");

            // Dynamic formula using INDEX/MATCH to retrieve the matching secondary category
            // The result will be a single cell (H1) that contains the matched category.
            Cell formulaCell = cells["H1"];
            formulaCell.SetDynamicArrayFormula(
                "=INDEX(C2:C5, MATCH(G1, C2:C5, 0))",
                new FormulaParseOptions(),
                true);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Primary series (values from B column)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Secondary series (values from D column) plotted on secondary axis
            chart.NSeries.Add("D2:D5", true);
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Set the secondary category axis data using the dynamic formula result
            chart.NSeries.SecondCategoryData = "H1";

            // Save the workbook
            workbook.Save("DynamicSecondaryAxis.xlsx");
        }
    }
}