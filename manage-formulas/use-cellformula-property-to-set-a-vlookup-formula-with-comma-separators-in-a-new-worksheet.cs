// Title: C# – Set a VLOOKUP formula with comma separators using Aspose.Cells
// Description: Shows how to create a new workbook, build a simple lookup table, add a lookup key, assign a VLOOKUP formula with comma separators via the Cell.Formula property, calculate all formulas, read the result, and save the file as VlookupDemo.xlsx.
// Keywords: Aspose.Cells | C# | VLOOKUP | Cell.Formula | comma separator | calculate formulas | save workbook | lookup table | Excel automation | locale formula separator
// Common Searches: Aspose.Cells set VLOOKUP formula with commas | C# VLOOKUP using Aspose.Cells Cell.Formula | how to calculate VLOOKUP in Aspose.Cells | save workbook after adding formula Aspose.Cells | locale specific formula separators Aspose.Cells
// Developer Intent: Insert a VLOOKUP formula that uses commas, evaluate it, and write the workbook to disk.
// Use Cases: Generate a price‑lookup sheet where item names are matched to prices via VLOOKUP. | Automate dynamic reports that require lookup calculations across many rows. | Create Excel files with pre‑calculated lookup results for downstream data pipelines.
// AI Prompts: Write C# code with Aspose.Cells that adds a VLOOKUP formula using comma separators, forces calculation, and saves the workbook. | Explain how to handle locale‑specific formula separators when setting formulas through Aspose.Cells. | Show how to apply the same VLOOKUP formula to a range of lookup values and retrieve each result programmatically.

using System;
using Aspose.Cells;

// Shows how to create a new workbook, build a simple lookup table, add a lookup key, assign a VLOOKUP formula with comma separators via the Cell.Formula property, calculate all formulas, read the result, and save the file as VlookupDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate a simple lookup table (A1:B4)
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(30);

        // The value we want to look up
        worksheet.Cells["D2"].PutValue("Banana");

        // Set VLOOKUP formula using comma separators
        // Syntax: =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        worksheet.Cells["E2"].Formula = "=VLOOKUP(D2, A2:B4, 2, FALSE)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the result of the VLOOKUP
        Console.WriteLine("VLOOKUP result: " + worksheet.Cells["E2"].Value);

        // Save the workbook to a file
        workbook.Save("VlookupDemo.xlsx");
    }
}
