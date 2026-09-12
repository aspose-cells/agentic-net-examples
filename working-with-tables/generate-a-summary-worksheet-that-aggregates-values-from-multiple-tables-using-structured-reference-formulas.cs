// Title: Generate an Excel summary sheet that totals amounts from multiple ListObject tables using structured reference formulas with Aspose.Cells for .NET
// AI Prompts: Create a workbook containing two worksheets, convert each data range into a ListObject named SalesQ1Table and SalesQ2Table, then add a third worksheet called Summary. | On the Summary sheet, assign B2 the formula =SUM(SalesQ1Table[Amount]), B3 the formula =SUM(SalesQ2Table[Amount]), and B4 the formula =SUM(SalesQ1Table[Amount], SalesQ2Table[Amount]). | Auto‑fit all columns in every worksheet and save the workbook as SummaryWorkbook.xlsx using Aspose.Cells.
// Common Searches: aspnet c# how to sum a column in an Aspose.Cells ListObject using structured references | create summary worksheet that aggregates data from multiple Excel tables with Aspose.Cells | Aspose.Cells example for grand total across two tables in .xlsx file | using ListObject tables and SUM formula in Aspose.Cells .NET
// Tags: structured reference SUM formula with Aspose.Cells ListObject | aggregate column values across multiple Excel tables in C# | create summary worksheet using Aspose.Cells tables | auto fit columns Aspose.Cells workbook | save workbook as .xlsx Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The program builds a workbook with two data sheets, converts each range into ListObject tables (SalesQ1Table, SalesQ2Table), adds a Summary sheet that uses structured reference formulas (=SUM(Table[Amount])) to compute quarterly totals and a grand total, auto‑fits all columns, and saves the file as SummaryWorkbook.xlsx via Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- First data worksheet ----------
            Worksheet ws1 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws1.Name = "SalesQ1";

            // Header
            ws1.Cells["A1"].PutValue("Product");
            ws1.Cells["B1"].PutValue("Amount");

            // Sample data
            ws1.Cells["A2"].PutValue("A");
            ws1.Cells["B2"].PutValue(1200);
            ws1.Cells["A3"].PutValue("B");
            ws1.Cells["B3"].PutValue(850);
            ws1.Cells["A4"].PutValue("C");
            ws1.Cells["B4"].PutValue(430);

            // Convert the range to a table (ListObject)
            int rows1 = ws1.Cells.MaxDataRow + 1;
            int cols1 = ws1.Cells.MaxDataColumn + 1;
            int tableIdx1 = ws1.ListObjects.Add(0, 0, rows1 - 1, cols1 - 1, true);
            ListObject table1 = ws1.ListObjects[tableIdx1];
            table1.DisplayName = "SalesQ1Table"; // Set table name

            // ---------- Second data worksheet ----------
            Worksheet ws2 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws2.Name = "SalesQ2";

            ws2.Cells["A1"].PutValue("Product");
            ws2.Cells["B1"].PutValue("Amount");

            ws2.Cells["A2"].PutValue("A");
            ws2.Cells["B2"].PutValue(1500);
            ws2.Cells["A3"].PutValue("B");
            ws2.Cells["B3"].PutValue(950);
            ws2.Cells["A4"].PutValue("C");
            ws2.Cells["B4"].PutValue(600);

            int rows2 = ws2.Cells.MaxDataRow + 1;
            int cols2 = ws2.Cells.MaxDataColumn + 1;
            int tableIdx2 = ws2.ListObjects.Add(0, 0, rows2 - 1, cols2 - 1, true);
            ListObject table2 = ws2.ListObjects[tableIdx2];
            table2.DisplayName = "SalesQ2Table"; // Set table name

            // ---------- Summary worksheet ----------
            Worksheet summary = workbook.Worksheets[workbook.Worksheets.Add()];
            summary.Name = "Summary";

            // Header for summary
            summary.Cells["A1"].PutValue("Quarter");
            summary.Cells["B1"].PutValue("Total Amount");

            // Q1 total using structured reference
            summary.Cells["A2"].PutValue("Q1");
            summary.Cells["B2"].Formula = "=SUM(SalesQ1Table[Amount])";

            // Q2 total using structured reference
            summary.Cells["A3"].PutValue("Q2");
            summary.Cells["B3"].Formula = "=SUM(SalesQ2Table[Amount])";

            // Grand total across both tables
            summary.Cells["A4"].PutValue("Grand Total");
            summary.Cells["B4"].Formula = "=SUM(SalesQ1Table[Amount], SalesQ2Table[Amount])";

            // Auto-fit columns for better readability
            ws1.AutoFitColumns();
            ws2.AutoFitColumns();
            summary.AutoFitColumns();

            // Determine output path and ensure directory exists
            string outputFile = "SummaryWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
