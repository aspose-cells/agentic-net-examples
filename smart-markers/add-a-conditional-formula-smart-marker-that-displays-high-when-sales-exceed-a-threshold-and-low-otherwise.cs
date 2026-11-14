// Title: Add a conditional smart marker in Aspose.Cells for .NET that returns “High” or “Low” based on a sales threshold
// Common Searches: Aspose.Cells C# conditional smart marker using IF formula | How to create a status column with smart markers that shows High or Low in Excel | WorkbookDesigner process smart marker with row‑relative IF condition in .NET | Smart marker expression for threshold evaluation of sales values in Aspose.Cells
// Tags: conditional smart marker IF formula Aspose.Cells | row‑relative smart marker expression C# | sales threshold evaluation smart marker | WorkbookDesigner conditional processing Excel | status column smart marker Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a Sales header and sample values, inserts a smart marker in column B that uses an IF formula to output "High" when the adjacent sales figure exceeds 1000 and "Low" otherwise, processes the markers with WorkbookDesigner, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SalesData";

        // Add header row
        sheet.Cells["A1"].PutValue("Sales");
        sheet.Cells["B1"].PutValue("Status");

        // Sample sales data
        double[] salesValues = { 800, 1200, 950, 1500, 700 };

        // Populate sales data starting from A2
        for (int i = 0; i < salesValues.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(salesValues[i]); // Column A (index 0)
        }

        // Insert a smart marker with a conditional formula in column B
        // The formula evaluates the sales value in column A of the same row.
        // If sales > 1000 => "High", else => "Low"
        // Smart marker syntax: &="IF(A2>1000,\"High\",\"Low\")"
        // The row number will be adjusted automatically during processing.
        sheet.Cells["B2"].PutValue("&=IF(A2>1000,\"High\",\"Low\")");

        // Apply smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.Process();

        // Save the result
        workbook.Save("ConditionalSmartMarker_Output.xlsx");
    }
}
