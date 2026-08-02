using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicSecondaryAxis
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate sample data
            // Primary categories (A column)
            cells["A1"].PutValue("Primary Category");
            cells["A2"].PutValue("Cat1");
            cells["A3"].PutValue("Cat2");
            cells["A4"].PutValue("Cat3");
            cells["A5"].PutValue("Cat4");

            // Secondary categories (B column) – this will be used for the secondary axis
            cells["B1"].PutValue("Secondary Category");
            cells["B2"].PutValue("Sub1");
            cells["B3"].PutValue("Sub2");
            cells["B4"].PutValue("Sub3");
            cells["B5"].PutValue("Sub4");

            // Values for the primary series (C column)
            cells["C1"].PutValue("Values");
            cells["C2"].PutValue(10);
            cells["C3"].PutValue(20);
            cells["C4"].PutValue(30);
            cells["C5"].PutValue(40);

            // 3. Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // 4. Set the primary series data (values)
            chart.NSeries.Add("C2:C5", true);

            // 5. Set primary category axis data
            chart.NSeries.CategoryData = "A2:A5";

            // 6. Build a dynamic range for the secondary category axis using INDEX/MATCH logic.
            //    The formula uses OFFSET together with MATCH to determine the number of rows
            //    that contain data in column B. The result spills into a range that we reference
            //    with the "#" operator.
            //    Formula: =OFFSET($B$2,0,0,MATCH(9.99999999999999E+307,$B$2:$B$5),1)
            //    MATCH finds the last numeric entry; for text we can use a large string match.
            //    Here the column contains text, so we use a trick with MATCH on a large string.
            string dynamicFormula = "=OFFSET($B$2,0,0,MATCH(\"~\", $B$2:$B$5, -1),1)";

            // Place the dynamic array formula in cell D2
            Cell dynamicCell = cells["D2"];
            dynamicCell.SetDynamicArrayFormula(dynamicFormula, new FormulaParseOptions(), true);

            // 7. Assign the spilled range (D2#) as the secondary category data source
            chart.NSeries.SecondCategoryData = "D2#";

            // 8. Plot the first series on the secondary value axis
            chart.NSeries[0].PlotOnSecondAxis = true;

            // 9. Optionally customize the secondary axis (e.g., title)
            Axis secondCategoryAxis = chart.SecondCategoryAxis;
            secondCategoryAxis.IsVisible = true;
            secondCategoryAxis.Title.Text = "Dynamic Secondary Categories";

            // 10. Save the workbook
            workbook.Save("DynamicSecondaryAxisDemo.xlsx");
        }
    }
}