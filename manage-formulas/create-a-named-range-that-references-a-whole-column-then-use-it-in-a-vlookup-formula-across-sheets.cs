// Title: C# – Define a Whole‑Column Named Range and Use It in a Cross‑Sheet VLOOKUP with Aspose.Cells
// Description: Creates a workbook, fills columns B and C on Sheet1, defines a named range that covers the entire B:C columns, adds Sheet2, inserts a VLOOKUP formula that references the named range, calculates formulas, outputs the result, and saves the file.
// Keywords: Aspose.Cells | C# | named range whole column | VLOOKUP | cross‑sheet lookup | Excel formula calculation | reference entire column | Workbook API
// Common Searches: Aspose.Cells define named range for whole column | C# VLOOKUP using named range across worksheets | How to reference entire column in Aspose.Cells formula | Calculate VLOOKUP after setting named range in Aspose.Cells
// Developer Intent: Create a column‑wide named range and reference it in a VLOOKUP formula on a different worksheet.
// Use Cases: Expose a B:C lookup table on Sheet1 as a named range and retrieve values from Sheet2. | Perform cross‑sheet data lookup without hard‑coding cell addresses. | Programmatically calculate formulas, read the VLOOKUP result, and save the workbook.
// AI Prompts: Show C# code that defines a whole‑column named range and uses it in a VLOOKUP on another sheet with Aspose.Cells. | Generate an Aspose.Cells example that populates sample data, creates a column‑wide named range, applies VLOOKUP across worksheets, calculates formulas, and saves the file. | Explain how to reference an entire column via a named range in Aspose.Cells and use it in a cross‑sheet VLOOKUP.

using System;
using Aspose.Cells;

// Creates a workbook, fills columns B and C on Sheet1, defines a named range that covers the entire B:C columns, adds Sheet2, inserts a VLOOKUP formula that references the named range, calculates formulas, outputs the result, and saves the file.
class NamedRangeVlookupDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Access the first worksheet and name it "Sheet1"
        Worksheet sheet1 = wb.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Populate sample data: column B contains keys, column C contains values
        sheet1.Cells["B1"].PutValue("Apple");
        sheet1.Cells["C1"].PutValue(100);
        sheet1.Cells["B2"].PutValue("Banana");
        sheet1.Cells["C2"].PutValue(200);
        sheet1.Cells["B3"].PutValue("Cherry");
        sheet1.Cells["C3"].PutValue(300);

        // Create a named range that refers to the whole columns B:C on Sheet1
        int nameIdx = wb.Worksheets.Names.Add("LookupTable");
        wb.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!$B:$C";

        // Add a second worksheet named "Sheet2"
        Worksheet sheet2 = wb.Worksheets.Add("Sheet2");

        // Use VLOOKUP on Sheet2 referencing the named range.
        // Lookup "Banana" and return the value from the second column of the table array.
        sheet2.Cells["A1"].Formula = "=VLOOKUP(\"Banana\", LookupTable, 2, FALSE)";

        // Calculate all formulas in the workbook
        wb.CalculateFormula();

        // Display the VLOOKUP result
        Console.WriteLine("VLOOKUP result: " + sheet2.Cells["A1"].Value);

        // Save the workbook to a file
        wb.Save("NamedRangeVlookupDemo.xlsx");
    }
}
