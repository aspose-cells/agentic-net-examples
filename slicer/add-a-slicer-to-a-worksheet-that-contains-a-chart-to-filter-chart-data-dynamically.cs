using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerChartDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (Fruit, Year, Amount)
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                string[] fruits = { "Apple", "Orange", "Banana", "Apple", "Orange", "Banana" };
                int[] years = { 2020, 2020, 2020, 2021, 2021, 2021 };
                int[] amounts = { 50, 70, 60, 80, 90, 100 };

                for (int i = 0; i < fruits.Length; i++)
                {
                    cells[i + 1, 0].PutValue(fruits[i]);   // Column A
                    cells[i + 1, 1].PutValue(years[i]);   // Column B
                    cells[i + 1, 2].PutValue(amounts[i]); // Column C
                }

                // Add a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("=Sheet1!A1:C7", "E2", "FruitPivot");
                PivotTable pivot = pivots[pivotIndex];

                // Configure pivot fields: Fruit as row, Year as column, Amount as data
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a column chart that uses the pivot table as its data source
                int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Convert the pivot table range (CellArea) to an A1‑style address string
                CellArea area = pivot.TableRange1;
                string startAddr = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                string endAddr = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                string rangeAddress = $"{sheet.Name}!{startAddr}:{endAddr}";

                // Use the address string for the chart series
                chart.NSeries.Add(rangeAddress, true);
                chart.Title.Text = "Fruit Sales by Year";

                // Add a slicer linked to the pivot table to filter by Fruit (placed at G2)
                int slicerIndex = sheet.Slicers.Add(pivot, "G2", "Fruit");
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Caption = "Fruit Filter";
                slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

                // Save the workbook
                string outputPath = "SlicerChartDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}