// Title: Add VLOOKUP formulas in a C# Aspose.Cells workbook to fetch quarterly target values and export to Excel
// AI Prompts: Generate C# code using Aspose.Cells that creates a worksheet, populates a source table with quarters and targets, and inserts a VLOOKUP formula in each cell of a result column to look up the target for a given quarter. | Write a C# loop with Aspose.Cells that assigns the formula VLOOKUP(D2, A$2:B$5, 2, FALSE) to cells E2 through E5, forces formula calculation, and saves the workbook as an .xlsx file. | Provide a complete Aspose.Cells example that demonstrates building a lookup list, applying VLOOKUP across a range, evaluating the formulas, and writing the final Excel file.
// Common Searches: how to programmatically add VLOOKUP to multiple cells with Aspose.Cells in C# | Aspose.Cells example for creating a lookup table and retrieving values using VLOOKUP | C# Aspose.Cells calculate formulas after inserting VLOOKUP and save workbook | using Aspose.Cells to generate a quarterly target report with VLOOKUP formulas | Aspose.Cells VLOOKUP syntax for Excel file generation in .NET
// Tags: Aspose.Cells VLOOKUP formula C# | C# generate Excel lookup range | Aspose.Cells calculate formulas | save Excel workbook with Aspose.Cells | quarterly target lookup Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, fills columns A and B with quarter and target data, sets up a lookup list in columns D and E, inserts VLOOKUP formulas in column E to retrieve targets based on quarter identifiers, forces formula evaluation, and saves the result as QuarterlyTargets.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data: Quarter and Target
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Target");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B4"].PutValue(2000);
        sheet.Cells["A5"].PutValue("Q4");
        sheet.Cells["B5"].PutValue(2500);

        // Define a range where we want to retrieve targets using VLOOKUP
        // List quarters in column D and place VLOOKUP results in column E
        sheet.Cells["D1"].PutValue("Quarter");
        sheet.Cells["E1"].PutValue("Target (Lookup)");
        sheet.Cells["D2"].PutValue("Q1");
        sheet.Cells["D3"].PutValue("Q2");
        sheet.Cells["D4"].PutValue("Q3");
        sheet.Cells["D5"].PutValue("Q4");

        // Apply VLOOKUP formula in column E for each quarter
        for (int row = 2; row <= 5; row++)
        {
            // VLOOKUP(lookup_value, table_array, col_index, FALSE)
            string formula = $"VLOOKUP(D{row},A$2:B$5,2,FALSE)";
            sheet.Cells[$"E{row}"].Formula = formula;
        }

        // Calculate formulas so that the workbook contains the evaluated values
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("QuarterlyTargets.xlsx");
    }
}
