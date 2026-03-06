using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableRefreshInfoDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 15;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 8;

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "D2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[ptIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column

            // Refresh the pivot table so that RefreshDate and RefreshedByWho are set
            pivot.RefreshData();
            pivot.CalculateData();

            // Output refresh information before saving
            Console.WriteLine("Before saving:");
            Console.WriteLine("Refresh Date : " + pivot.RefreshDate);
            Console.WriteLine("Refreshed By : " + pivot.RefreshedByWho);

            // Save the workbook to an XLSX file
            string filePath = "PivotRefreshInfoDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Reload the workbook to demonstrate persistence of the properties
            Workbook loadedWb = new Workbook(filePath);
            PivotTable loadedPivot = loadedWb.Worksheets[0].PivotTables[0];

            // Output refresh information after reloading
            Console.WriteLine("\nAfter reloading:");
            Console.WriteLine("Refresh Date : " + loadedPivot.RefreshDate);
            Console.WriteLine("Refreshed By : " + loadedPivot.RefreshedByWho);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableRefreshInfoDemo.Run();
        }
    }
}