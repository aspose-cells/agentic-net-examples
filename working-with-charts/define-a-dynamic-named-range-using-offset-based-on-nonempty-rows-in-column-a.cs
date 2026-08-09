// Title: Aspose.Cells for .NET – Create a Dynamic Named Range with OFFSET & COUNTA (Column A)
// Description: C# example that builds a workbook, inserts sample data in column A, adds a named range called DynamicRange using the OFFSET formula combined with COUNTA to automatically cover all non‑empty rows, resolves the range to get its address, and saves the file as DynamicNamedRange.xlsx. Ideal for charts and formulas that need a self‑adjusting data source.
// Keywords: Aspose.Cells dynamic named range | OFFSET function .NET | COUNTA Excel formula | C# Excel named range example | auto‑expanding range Aspose | chart data source dynamic range | programmatic Excel range | Aspose.Cells workbook sample | Excel OFFSET COUNTA code
// Common Searches: Aspose.Cells define dynamic named range with OFFSET | C# OFFSET COUNTA named range Aspose.Cells | auto expanding range column A Aspose.Cells | how to create dynamic range for chart in .NET | Aspose.Cells named range that grows with data
// Developer Intent: Programmatically add a named range that automatically expands to include every non‑empty cell in column A.
// Use Cases: Link a chart series to a range that grows as new rows are added. | Reference a continuously updating dataset in formulas without manual range changes. | Export workbooks where the data source size varies per user or per run.
// AI Prompts: Modify the OFFSET formula to start at A2 while keeping the dynamic height. | Add code that appends rows to column A and then prints the updated DynamicRange address. | Show how to set DynamicRange as the source series for a chart created with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that builds a workbook, inserts sample data in column A, adds a named range called DynamicRange using the OFFSET formula combined with COUNTA to automatically cover all non‑empty rows, resolves the range to get its address, and saves the file as DynamicNamedRange.xlsx. Ideal for charts and formulas that need a self‑adjusting data source.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Sample data in column A (optional, just for demonstration)
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);

            // Add a named range that expands automatically with non‑empty rows in column A
            int nameIndex = wb.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = wb.Worksheets.Names[nameIndex];
            // OFFSET(start, rows, cols, height, width)
            // Height is determined by COUNTA of column A (non‑empty cells)
            dynamicName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

            // Resolve the named range to an Aspose.Cells.Range object
            Aspose.Cells.Range resolvedRange = dynamicName.GetRange();
            Console.WriteLine("Dynamic range address: " + resolvedRange.Address);

            // Save the workbook
            wb.Save("DynamicNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
