using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ShowReportFilterPagesDemo
    {
        public static void Run()
        {
            // Load an existing workbook that contains a pivot table (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
                return;

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
                return;

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure the pivot table has at least one page field before calling the method
            if (pivotTable.PageFields.Count > 0)
            {
                // Show report filter pages for each page field
                foreach (PivotField pageField in pivotTable.PageFields)
                {
                    pivotTable.ShowReportFilterPage(pageField);
                }
            }

            // Save the workbook with the changes (XLSX format)
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShowReportFilterPagesDemo.Run();
        }
    }
}