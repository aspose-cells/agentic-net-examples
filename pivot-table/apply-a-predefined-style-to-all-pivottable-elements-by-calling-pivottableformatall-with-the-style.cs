using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ApplyStyleToAllPivotTables
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data and a pivot table (optional setup)
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 5;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 7;

            int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // -------------------------------------------------
            // Define the predefined style to be applied
            // -------------------------------------------------
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // -------------------------------------------------
            // Apply the style to every PivotTable in the workbook
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    pt.FormatAll(style);
                }
            }

            // -------------------------------------------------
            // Save the workbook with the applied styles
            // -------------------------------------------------
            string outputPath = "AllPivotTablesStyled.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}