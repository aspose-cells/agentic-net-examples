// Title: Create a dashboard worksheet with a SUM formula that aggregates data from multiple sheets using Aspose.Cells for .NET
// Common Searches: how to add a SUM formula that references multiple worksheets in Aspose.Cells C# | Aspose.Cells .NET create dashboard sheet that totals ranges across sheets | C# code to sum A1:A10 from Sheet1, Sheet2, and Sheet3 using Aspose.Cells | generate an Excel workbook with aggregated totals on a dashboard worksheet using Aspose.Cells | populate several worksheets and compute a combined sum in a single cell with Aspose.Cells
// Tags: Aspose.Cells SUM formula across worksheets | C# create dashboard worksheet Aspose.Cells | Excel multi-sheet aggregation Aspose.Cells | programmatic workbook generation Aspose.Cells .NET | populate cells and apply formula Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, adds three worksheets named Sheet1, Sheet2, and Sheet3, fills cells A1:A10 in each sheet with incremental values, then adds a Dashboard worksheet where cell A1 contains a SUM formula that references the same range on all three sheets, and finally saves the workbook as Dashboard.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook wb = new Workbook();

            // Get reference to the first worksheet and rename it
            Worksheet ws1 = wb.Worksheets[0];
            ws1.Name = "Sheet1";

            // Add second worksheet and rename
            int ws2Idx = wb.Worksheets.Add();
            Worksheet ws2 = wb.Worksheets[ws2Idx];
            ws2.Name = "Sheet2";

            // Add third worksheet and rename
            int ws3Idx = wb.Worksheets.Add();
            Worksheet ws3 = wb.Worksheets[ws3Idx];
            ws3.Name = "Sheet3";

            // Populate sample data in each sheet (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                ws1.Cells[i, 0].PutValue(i + 1);               // 1..10
                ws2.Cells[i, 0].PutValue((i + 1) * 2);         // 2..20 (step 2)
                ws3.Cells[i, 0].PutValue((i + 1) * 3);         // 3..30 (step 3)
            }

            // Add a dashboard sheet and display the aggregated result using a direct SUM formula
            int dashIdx = wb.Worksheets.Add();
            Worksheet dashboard = wb.Worksheets[dashIdx];
            dashboard.Name = "Dashboard";
            dashboard.Cells["A1"].Formula = "=SUM(Sheet1!A1:A10,Sheet2!A1:A10,Sheet3!A1:A10)";

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Dashboard.xlsx");

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
