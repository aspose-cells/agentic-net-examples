// Title: Consolidate Data from Multiple Worksheets into a Pivot Table Using Named Ranges – Aspose.Cells for .NET
// Description: This example creates a workbook with two sheets, defines named ranges (Data1, Data2) that point to each sheet's A1:B4 area, and builds a single pivot table on Sheet1 using those ranges as the source. The pivot groups by "Category" and sums "Value", then saves the file as ConsolidatedPivot.xlsx.
// Keywords: Aspose.Cells | C# | named range | pivot table | multiple worksheets | data consolidation | Excel automation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells named range pivot source | consolidate worksheets into one pivot table .NET | use named ranges for pivot table Aspose | pivot table from multiple sheets Aspose.Cells | C# Aspose.Cells data consolidation example
// Developer Intent: Generate a pivot table that pulls data from several worksheets by referencing named ranges.
// Use Cases: Merge regional sales sheets into a single pivot report using named ranges. | Create a financial summary that aggregates quarterly figures stored on separate tabs. | Build an inventory dashboard that consolidates stock counts from multiple department sheets.
// AI Prompts: Add more named ranges to the source array for the consolidated pivot table in Aspose.Cells. | Change the pivot calculation from Sum to Average for the "Value" field using Aspose.Cells. | Show how to refresh the consolidated pivot after modifying the source worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook with two sheets, defines named ranges (Data1, Data2) that point to each sheet's A1:B4 area, and builds a single pivot table on Sheet1 using those ranges as the source. The pivot groups by "Category" and sums "Value", then saves the file as ConsolidatedPivot.xlsx.
class ConsolidateUsingNamedRanges
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // ---------- Worksheet 1 ----------
            Worksheet ws1 = wb.Worksheets[0];
            ws1.Name = "Sheet1";

            // Populate data in Sheet1
            ws1.Cells["A1"].PutValue("Category");
            ws1.Cells["B1"].PutValue("Value");
            ws1.Cells["A2"].PutValue("A");
            ws1.Cells["B2"].PutValue(10);
            ws1.Cells["A3"].PutValue("B");
            ws1.Cells["B3"].PutValue(20);
            ws1.Cells["A4"].PutValue("C");
            ws1.Cells["B4"].PutValue(30);

            // ---------- Worksheet 2 ----------
            Worksheet ws2 = wb.Worksheets.Add("Sheet2");

            // Populate data in Sheet2
            ws2.Cells["A1"].PutValue("Category");
            ws2.Cells["B1"].PutValue("Value");
            ws2.Cells["A2"].PutValue("A");
            ws2.Cells["B2"].PutValue(15);
            ws2.Cells["A3"].PutValue("B");
            ws2.Cells["B3"].PutValue(25);
            ws2.Cells["A4"].PutValue("C");
            ws2.Cells["B4"].PutValue(35);

            // ---------- Create Named Ranges ----------
            // Named range "Data1" refers to Sheet1!A1:B4
            int nameIdx1 = wb.Worksheets.Names.Add("Data1");
            wb.Worksheets.Names[nameIdx1].RefersTo = "=Sheet1!$A$1:$B$4";

            // Named range "Data2" refers to Sheet2!A1:B4
            int nameIdx2 = wb.Worksheets.Names.Add("Data2");
            wb.Worksheets.Names[nameIdx2].RefersTo = "=Sheet2!$A$1:$B$4";

            // Retrieve the address strings (without the leading '=') for use as source data
            string range1 = wb.Worksheets.Names["Data1"].RefersTo.TrimStart('=');
            string range2 = wb.Worksheets.Names["Data2"].RefersTo.TrimStart('=');
            string[] sourceData = new string[] { range1, range2 };

            // ---------- Create Pivot Table Using Consolidated Ranges ----------
            // Destination cell D3 corresponds to row index 2, column index 3 (zero‑based)
            PivotTableCollection pivots = ws1.PivotTables;
            int pivotIdx = pivots.Add(sourceData, false, null, 2, 3, "ConsolidatedPivot");
            PivotTable pivot = pivots[pivotIdx];

            // Configure pivot fields: Category as row field, Value as data field (sum)
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // ---------- Save Workbook ----------
            wb.Save("ConsolidatedPivot.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
