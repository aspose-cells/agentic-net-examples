// Title: Generate a column chart with data labels linked to a dynamic named range and automatically resize the chart shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bind chart series to a structured named range that expands with a ListObject | C# resize Aspose.Cells chart shape after adding rows to an Excel table | How to show category names and values in data labels of an Aspose.Cells column chart | Create dynamic named ranges for chart categories and values using Aspose.Cells tables | Update chart range automatically when ListObject size changes in Aspose.Cells
// Tags: dynamic named range chart series Aspose.Cells | auto resize chart shape Aspose.Cells | structured references chart Aspose.Cells | listobject table expansion chart update | data labels show value and category Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, populates categories and values, converts the data into a ListObject (Excel table) and defines structured named ranges that reference the table columns. A column chart is added using these named ranges, with data labels configured to display both the value and the category name. After adding new rows, the ListObject is resized, causing the named ranges to expand automatically, and the chart shape's dimensions are recalculated based on the updated table size. The workbook is then saved as DynamicRangeChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate initial data (Category / Value)
            // ------------------------------------------------------------
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            string[] initialCategories = { "A", "B", "C" };
            double[] initialValues = { 10, 20, 30 };

            for (int i = 0; i < initialCategories.Length; i++)
            {
                cells[i + 1, 0].PutValue(initialCategories[i]); // Column A
                cells[i + 1, 1].PutValue(initialValues[i]);    // Column B
            }

            // ------------------------------------------------------------
            // 2. Create a ListObject (Excel Table) – this will act as a
            //    dynamic named range because the table can be resized.
            // ------------------------------------------------------------
            ListObjectCollection tables = worksheet.ListObjects;
            int tableIdx = tables.Add(0, 0, initialCategories.Length, 1, true);
            ListObject table = tables[tableIdx];
            table.DisplayName = "DataTable";
            table.ShowHeaderRow = true;
            table.ShowTotals = false;

            // ------------------------------------------------------------
            // 3. Define structured named ranges that point to the table columns.
            //    Structured references automatically expand when the table grows.
            // ------------------------------------------------------------
            int valuesIdx = workbook.Worksheets.Names.Add("ValuesRange");
            Name valuesRange = workbook.Worksheets.Names[valuesIdx];
            valuesRange.RefersTo = "=DataTable[Value]";

            int categoriesIdx = workbook.Worksheets.Names.Add("CategoriesRange");
            Name categoriesRange = workbook.Worksheets.Names[categoriesIdx];
            categoriesRange.RefersTo = "=DataTable[Category]";

            // ------------------------------------------------------------
            // 4. Add a chart that uses the named ranges.
            // ------------------------------------------------------------
            int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 25, 8);
            Chart chart = worksheet.Charts[chartIdx];

            // Add series using the named range for values
            chart.NSeries.Add("=ValuesRange", true);

            // Set category data using the structured range
            // Note: In some older Aspose.Cells versions the Series class may not expose CategoryData.
            // If available, uncomment the following line:
            // chart.NSeries[0].CategoryData = "=CategoriesRange";

            // Enable data labels to show the values and categories
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowCategoryName = true;

            // Keep a reference to the ChartShape – we will resize it later.
            ChartShape chartShape = (ChartShape)chart.ChartObject;

            // ------------------------------------------------------------
            // 5. Simulate adding more rows to the data table.
            //    After adding rows we resize the ListObject so the structured
            //    named ranges automatically include the new data.
            // ------------------------------------------------------------
            string[] newCategories = { "D", "E" };
            double[] newValues = { 40, 50 };

            int startRowForNewData = table.EndRow + 1; // first empty row after the table
            for (int i = 0; i < newCategories.Length; i++)
            {
                cells[startRowForNewData + i, 0].PutValue(newCategories[i]); // Column A
                cells[startRowForNewData + i, 1].PutValue(newValues[i]);    // Column B
            }

            // Resize the table to include the newly added rows
            int newEndRow = startRowForNewData + newCategories.Length - 1;
            table.Resize(table.StartRow, table.StartColumn, newEndRow, table.EndColumn, true);

            // ------------------------------------------------------------
            // 6. Resize the chart shape based on the new size of the table.
            //    This is a simple example: height grows with row count,
            //    width grows with column count.
            // ------------------------------------------------------------
            int rowCount = newEndRow - table.StartRow + 1;          // total rows in the table (including header)
            int columnCount = table.EndColumn - table.StartColumn + 1;

            const int pixelsPerRow = 15;
            const int pixelsPerColumn = 80;

            chartShape.Height = (rowCount + 2) * pixelsPerRow; // +2 for extra space
            chartShape.Width = columnCount * pixelsPerColumn;

            // ------------------------------------------------------------
            // 7. Recalculate any formulas (good practice)
            // ------------------------------------------------------------
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 8. Save the workbook
            // ------------------------------------------------------------
            string outputPath = "DynamicRangeChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
