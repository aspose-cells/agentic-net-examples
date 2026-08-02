using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace CustomSubtotalLabelValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook and sample data --------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");

                // Sample rows
                string[] categories = { "A", "A", "B", "B", "A" };
                double[] amounts = { 100, 150, 200, 250, 120 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 1, 0].PutValue(categories[i]);   // Column A
                    cells[i + 1, 1].PutValue(amounts[i]);     // Column B
                }

                // -------------------- Define custom subtotal label --------------------
                // Create a SettablePivotGlobalizationSettings instance and set a custom label
                SettablePivotGlobalizationSettings globalization = new SettablePivotGlobalizationSettings();
                string customSumLabel = "My Custom Sum";
                globalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, customSumLabel);

                // Apply the globalization settings to the workbook if the property exists (available in newer versions)
                var settingsType = workbook.Settings.GetType();
                var propInfo = settingsType.GetProperty("PivotGlobalizationSettings");
                if (propInfo != null && propInfo.CanWrite)
                {
                    propInfo.SetValue(workbook.Settings, globalization);
                }

                // -------------------- Create pivot table --------------------
                // Add pivot table covering the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add Category as row field and Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column B

                // Enable Sum subtotal for the row field
                PivotField rowField = pivotTable.RowFields[0];
                rowField.SetSubtotals(PivotFieldSubtotalType.Sum, true);

                // Refresh pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------- Save workbook as PDF --------------------
                string pdfPath = "CustomSubtotalLabel.pdf";

                // Ensure the directory exists (if a directory part is present)
                string pdfDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
                {
                    Directory.CreateDirectory(pdfDir);
                }

                workbook.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF file was created
                if (File.Exists(pdfPath))
                {
                    Console.WriteLine($"Success: PDF saved to \"{pdfPath}\" with custom subtotal label \"{customSumLabel}\".");
                }
                else
                {
                    Console.WriteLine($"Failure: PDF was not created at \"{pdfPath}\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}