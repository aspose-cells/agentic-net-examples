// Title: Filter rows with AutoFilter and compute MINIFS on visible data using Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that applies an AutoFilter on a worksheet, hides rows, then calls Workbook.CalculateFormula so that a MINIFS formula evaluates only the visible rows in Aspose.Cells. | Show how to change the MINIFS criteria value while keeping the AutoFilter active and retrieve the updated result with Aspose.Cells for .NET. | Explain how to verify which rows remain visible after filtering and how to access the MINIFS result programmatically in Aspose.Cells.
// Common Searches: Aspose.Cells C# calculate MINIFS after applying AutoFilter to a column | How to make MINIFS consider only filtered rows in Aspose.Cells .NET | C# example of using Worksheet.AutoFilter with MINIFS formula in Aspose.Cells | Retrieve MINIFS result from visible rows after Workbook.CalculateFormula in Aspose.Cells | Aspose.Cells filter data then evaluate formulas for visible cells
// Tags: auto-filter worksheet Aspose.Cells C# | MINIFS formula with filtered rows .NET | calculate formulas after applying filter Aspose.Cells | visible rows MINIFS Aspose.Cells | Aspose.Cells workbook.CalculateFormula usage

using System;
using Aspose.Cells;

namespace AsposeCellsMinifsFilterDemo
{
    // Demonstrates applying an AutoFilter on column B to show only category "X", then calculating a MINIFS formula that finds the minimum value in column A for the visible rows, using Aspose.Cells for .NET (C#).
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Column A: numeric values
            // Column B: categories used for filtering
            cells["A1"].PutValue("Value");
            cells["B1"].PutValue("Category");
            cells["A2"].PutValue(10); cells["B2"].PutValue("X");
            cells["A3"].PutValue(20); cells["B3"].PutValue("Y");
            cells["A4"].PutValue(5);  cells["B4"].PutValue("X");
            cells["A5"].PutValue(30); cells["B5"].PutValue("Y");
            cells["A6"].PutValue(15); cells["B6"].PutValue("X");

            // Insert MINIFS formula that finds the minimum value in column A
            // where the corresponding category in column B equals "X"
            cells["C1"].PutValue("MinIfsResult");
            cells["C2"].Formula = "=MINIFS(A2:A6, B2:B6, \"X\")";

            // Apply an auto‑filter on the header row (A1:B1)
            sheet.AutoFilter.Range = "A1:B6";

            // Filter column B to show only rows with category "X"
            sheet.AutoFilter.Filter(1, "X"); // fieldIndex 1 corresponds to column B
            sheet.AutoFilter.Refresh(); // Apply the filter (hides rows with "Y")

            // Calculate all formulas after the filter is applied
            workbook.CalculateFormula();

            // Retrieve and display the result of the MINIFS formula
            Console.WriteLine("MINIFS result (should consider only visible rows with category \"X\"):");
            Console.WriteLine("C2 = " + cells["C2"].Value);

            // For verification, also output which rows are visible
            Console.WriteLine("\nVisible rows after filtering:");
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                if (!sheet.Cells.Rows[row].IsHidden)
                {
                    Console.WriteLine($"Row {row + 1}: Value={cells[row, 0].Value}, Category={cells[row, 1].Value}");
                }
            }

            // Save the workbook (optional, demonstrates usage of the save rule)
            workbook.Save("MinifsFilterDemo.xlsx");
        }
    }
}
