using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsTimelineToPdf
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate worksheet with sample sales data (Date and Sales)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["B2"].PutValue(1500);
                sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["B3"].PutValue(2300);
                sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["B4"].PutValue(1800);
                sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
                sheet.Cells["B5"].PutValue(2100);

                // 3. Add a pivot table based on the sales data (placed at D2)
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot.RefreshData();
                pivot.CalculateData();

                // 4. Insert a Timeline control linked to the pivot table (starting at G2)
                try
                {
                    sheet.Timelines.Add(pivot, 1, 6, "Date");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Timeline could not be added: {ex.Message}");
                }

                // 5. Save the workbook to a temporary Excel file
                string excelPath = "SalesTimeline.xlsx";
                workbook.Save(excelPath);

                // 6. Convert the saved Excel file to PDF using ConversionUtility
                string pdfPath = "SalesTimeline.pdf";
                if (File.Exists(excelPath))
                {
                    ConversionUtility.Convert(excelPath, pdfPath);
                    Console.WriteLine($"Workbook saved to '{excelPath}' and converted to PDF at '{pdfPath}'.");
                }
                else
                {
                    Console.WriteLine($"Excel file '{excelPath}' not found. Conversion skipped.");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred: {e.Message}");
            }
        }
    }
}