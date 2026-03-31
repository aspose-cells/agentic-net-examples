using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOverview
{
    class Program
    {
        static void Main()
        {
            // Apply license if it exists
            var license = new License();
            const string licensePath = "Aspose.Cells.NET.lic";
            if (File.Exists(licensePath))
            {
                license.SetLicense(licensePath);
            }

            // Core object – Workbook
            var workbook = new Workbook();

            // Work with worksheets and cells
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Write header
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");

            // Write sample rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(85);

            // Use CellsHelper for utility operations
            string colName = CellsHelper.ColumnIndexToName(2); // "C"
            int colIndex = CellsHelper.ColumnNameToIndex(colName); // 2

            string unsafeName = "My/Invalid:Sheet*Name?";
            string safeName = CellsHelper.CreateSafeSheetName(unsafeName, '_');

            // Set built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "Jane Doe";
            workbook.BuiltInDocumentProperties.Title = "Sales Report";
            workbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // Scenario manager
            var scenarios = sheet.Scenarios;
            int baseIdx = scenarios.Add("Base");
            scenarios.ActiveIndex = baseIdx;
            scenarios[baseIdx].Comment = "Base scenario";

            // Save in multiple formats
            workbook.Save("Overview.xlsx");
            workbook.Save("Overview.pdf");
            workbook.Save("Overview.csv");

            // Console overview
            Console.WriteLine("Aspose.Cells for .NET – Core Capabilities:");
            Console.WriteLine("- Create, load, modify, and save Excel workbooks (XLS, XLSX, CSV, PDF, etc.)");
            Console.WriteLine("- Rich API for cells, formulas, styles, charts, tables, pivot tables");
            Console.WriteLine("- Helper utilities via CellsHelper (address conversion, safe names, DPI, etc.)");
            Console.WriteLine("- Document properties, custom XML, VBA/macros handling");
            Console.WriteLine("- Scenario manager for what‑if analysis");
            Console.WriteLine("- Cloud‑ready settings (IsCloudPlatform, StartupPath)");
            Console.WriteLine("- Metered licensing support via Aspose.Cells.Metered");
            Console.WriteLine("- AI integration via Aspose.Cells.AI (CellsAI) for natural‑language queries");
            Console.WriteLine("Generated files: Overview.xlsx, Overview.pdf, Overview.csv");
        }
    }
}