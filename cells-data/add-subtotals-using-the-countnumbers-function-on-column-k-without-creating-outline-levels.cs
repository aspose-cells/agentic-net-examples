// Title: Aspose.Cells .NET – Add CountNums Subtotal to Column K without Outline Levels
// Description: C# example that creates a workbook, writes a header, three numeric values and a text entry into column K, defines a CellArea for rows 0‑4, and calls Worksheet.Cells.Subtotal with ConsolidationFunction.CountNums to insert a numeric‑count subtotal in the same column. No outline hierarchy is generated and the file is saved as SubtotalCountNumsColumnK.xlsx.
// Keywords: Aspose.Cells subtotal CountNums C# | Worksheet.Cells.Subtotal example | count numeric cells Aspose | Excel subtotal without outline | column K subtotal Aspose.Cells | C# Excel automation CountNums
// Common Searches: Aspose.Cells CountNums subtotal column K | C# add subtotal without outline levels | How to count numeric cells with Aspose.Cells | Worksheet.Cells.Subtotal usage example | Excel subtotal CountNums C# code
// Developer Intent: Insert a CountNums subtotal for column K while preventing the creation of outline levels.
// Use Cases: Produce a summary row that counts only numeric entries in a data column for financial reports. | Generate Excel files where text values are ignored in subtotal calculations, simplifying downstream analysis. | Create clean worksheets with subtotal rows that do not introduce collapsible outline groups.
// AI Prompts: Write C# code using Aspose.Cells to apply a CountNums subtotal on column K without creating outline levels. | Show how to call Worksheet.Cells.Subtotal with ConsolidationFunction.CountNums, grouping by the same column and adding the subtotal to that column. | Explain why the CountNums function ignores non‑numeric cells when generating subtotals in Aspose.Cells.

using Aspose.Cells;

// C# example that creates a workbook, writes a header, three numeric values and a text entry into column K, defines a CellArea for rows 0‑4, and calls Worksheet.Cells.Subtotal with ConsolidationFunction.CountNums to insert a numeric‑count subtotal in the same column. No outline hierarchy is generated and the file is saved as SubtotalCountNumsColumnK.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in column K (zero‑based index 10)
        // Header
        cells[0, 10].PutValue("Numbers");
        // Numeric values
        cells[1, 10].PutValue(5);
        cells[2, 10].PutValue(10);
        cells[3, 10].PutValue(15);
        // Non‑numeric value (should be ignored by CountNums)
        cells[4, 10].PutValue("Text");

        // Define the range that includes the header and data (A1 is row 0, column 10)
        CellArea area = CellArea.CreateCellArea(0, 10, 4, 10);

        // Apply subtotals using the CountNums function on column K.
        // groupBy = 0 (group by the first column of the area, which is column K itself)
        // totalList = new int[] { 0 } (add subtotal for the same column)
        worksheet.Cells.Subtotal(area, 0, ConsolidationFunction.CountNums, new int[] { 0 });

        // Save the workbook
        workbook.Save("SubtotalCountNumsColumnK.xlsx");
    }
}
