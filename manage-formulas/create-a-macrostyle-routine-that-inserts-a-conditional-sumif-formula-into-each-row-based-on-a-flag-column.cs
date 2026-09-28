// Title: Insert a conditional SUMIF formula into each data row using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to determine the data range, then writes a SUMIF formula into column D for every row, referencing a flag column in A and values in B. | Create a macro‑style routine with Aspose.Cells that loops through rows, builds a dynamic SUMIF expression based on the current row index, assigns it to the cell, and saves the workbook.
// Common Searches: Aspose.Cells C# add SUMIF formula to each row based on flag column | How to programmatically set conditional SUMIF in Excel using Aspose.Cells .NET | C# loop through worksheet rows and insert dynamic SUMIF range with Aspose.Cells | Generate Excel file with conditional aggregation formula using Aspose.Cells for .NET | Save workbook after inserting formulas with Aspose.Cells C# example
// Tags: Aspose.Cells insert SUMIF formula | C# loop assign Excel formulas | dynamic range calculation Aspose.Cells | conditional aggregation based on flag column | save workbook with Aspose.Cells .NET

using Aspose.Cells;

// The example creates a workbook, optionally fills sample flag and value columns, calculates the data range, and inserts a SUMIF formula into column D for each row that sums values from column B where the flag in column A matches the current row, then saves the file as ConditionalSumIf.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Sample data setup (optional, can be removed)
        // Column A: Flag column (e.g., "Yes"/"No")
        // Column B: Values to be summed
        // -------------------------------------------------
        sheet.Cells["A2"].PutValue("Yes");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("No");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Yes");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue("No");
        sheet.Cells["B5"].PutValue(40);
        // -------------------------------------------------

        // Determine the range of data rows
        int firstDataRow = 1; // zero‑based index for row 2 (header assumed at row 1)
        int lastRow = sheet.Cells.MaxDataRow; // last row containing data

        // If there is no data, define a default range (adjust as needed)
        if (lastRow < firstDataRow)
            lastRow = firstDataRow + 10; // placeholder range

        // Insert a conditional SUMIF formula into column D for each data row
        // Formula: =SUMIF($A$2:$A${lastRow+1}, A{currentRow+1}, $B$2:$B${lastRow+1})
        for (int row = firstDataRow; row <= lastRow; row++)
        {
            string formula = $"SUMIF($A${firstDataRow + 1}:$A${lastRow + 1}, A{row + 1}, $B${firstDataRow + 1}:$B${lastRow + 1})";
            sheet.Cells[row, 3].Formula = formula; // Column D (index 3)
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("ConditionalSumIf.xlsx");
    }
}
