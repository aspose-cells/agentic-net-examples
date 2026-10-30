// Title: Insert subtotal rows above each group with summaryPosition set to top using Aspose.Cells for .NET
// Common Searches: how to add a subtotal row before each group in Excel with Aspose.Cells C# | Aspose.Cells example for inserting summary rows at the top of grouped data | C# calculate group totals and place them above the data using Aspose.Cells | using summaryPosition top to create subtotals for each category in Aspose.Cells | insert dynamic subtotal rows before groups in an Excel workbook with Aspose.Cells .NET
// Tags: insert subtotal rows before group Aspose.Cells | summaryPosition top Excel Aspose.Cells | grouped data aggregation C# Aspose.Cells | dynamic row insertion Excel .NET | calculate group totals Aspose.Cells | Excel workbook subtotal example C#

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

// Demonstrates creating a workbook, populating group and amount columns, computing sums per group, inserting a subtotal row above each group's first row by using summaryPosition set to top, and saving the file as Subtotals_Top.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Group column in A, Amount column in B)
            sheet.Cells["A1"].PutValue("Group");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("East");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("East");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("West");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("West");
            sheet.Cells["B5"].PutValue(200);
            sheet.Cells["A6"].PutValue("North");
            sheet.Cells["B6"].PutValue(90);
            sheet.Cells["A7"].PutValue("North");
            sheet.Cells["B7"].PutValue(110);

            // Calculate subtotals manually (summary rows before each group)
            var groupFirstRow = new Dictionary<string, int>();
            var groupSum = new Dictionary<string, double>();

            // Data rows start at index 1 (row 2 in Excel)
            for (int i = 1; i <= 6; i++)
            {
                string group = sheet.Cells[i, 0].StringValue;
                double amount = sheet.Cells[i, 1].DoubleValue;

                if (!groupFirstRow.ContainsKey(group))
                    groupFirstRow[group] = i; // store first occurrence row index

                if (!groupSum.ContainsKey(group))
                    groupSum[group] = 0;

                groupSum[group] += amount;
            }

            // Insert subtotal rows before each group's first row
            var orderedGroups = groupFirstRow.OrderBy(kv => kv.Value).ToList();
            int offset = 0; // tracks rows inserted so far

            foreach (var kv in orderedGroups)
            {
                int insertRow = kv.Value + offset;
                sheet.Cells.InsertRows(insertRow, 1);
                sheet.Cells[insertRow, 0].PutValue($"{kv.Key} Total");
                sheet.Cells[insertRow, 1].PutValue(groupSum[kv.Key]);
                offset++; // one row added
            }

            // Save the workbook
            string outputPath = "Subtotals_Top.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
