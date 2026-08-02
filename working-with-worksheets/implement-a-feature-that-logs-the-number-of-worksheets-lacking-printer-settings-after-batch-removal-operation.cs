using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            Worksheet ws1 = workbook.Worksheets[0]; // default sheet
            ws1.Name = "SheetWithPrinter";

            Worksheet ws2 = workbook.Worksheets.Add("SheetWithoutPrinter1");
            Worksheet ws3 = workbook.Worksheets.Add("SheetWithoutPrinter2");
            Worksheet ws4 = workbook.Worksheets.Add("SheetWithPrinter2");

            // Set printer settings for some worksheets (simulated with a non‑empty byte array)
            ws1.PageSetup.PrinterSettings = new byte[] { 1, 2, 3 };
            ws4.PageSetup.PrinterSettings = new byte[] { 4, 5, 6 };
            // ws2 and ws3 intentionally leave PrinterSettings as null (lacking printer settings)

            // Batch removal: remove the second worksheet by index (example operation)
            // This demonstrates the removal logic; more removals can be added as needed.
            workbook.Worksheets.RemoveAt(1); // removes "SheetWithoutPrinter1"

            // After removal, count worksheets that do NOT have printer settings
            int lackingPrinterSettingsCount = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // PrinterSettings is a byte[]; consider null or empty as lacking settings
                if (ws.PageSetup.PrinterSettings == null || ws.PageSetup.PrinterSettings.Length == 0)
                {
                    lackingPrinterSettingsCount++;
                }
            }

            // Log the result
            Console.WriteLine($"Number of worksheets lacking printer settings: {lackingPrinterSettingsCount}");

            // Save the workbook (lifecycle save)
            workbook.Save("ResultWorkbook.xlsx");
        }
    }
}