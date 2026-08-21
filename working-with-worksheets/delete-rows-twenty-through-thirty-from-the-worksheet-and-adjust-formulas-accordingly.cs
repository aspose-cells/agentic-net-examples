// Title: C# – Delete rows 20‑30 in an Excel sheet and auto‑update formulas with Aspose.Cells
// Description: Load a workbook, remove rows 20‑30 (zero‑based index 19, count 11) using Cells.DeleteRows, set the third parameter to true so all formula references adjust automatically, then save the file. Includes a ready‑to‑run Aspose.Cells .NET example.
// Keywords: Aspose.Cells DeleteRows C# | remove multiple rows Excel .NET | auto adjust formulas Aspose | Excel row deletion example | C# Aspose.Cells GitHub | Excel automation US developers | Excel automation UK developers | Excel automation India developers
// Common Searches: How to delete rows 20 to 30 with Aspose.Cells C# | Aspose.Cells delete rows and keep formulas correct | C# code to remove a range of rows in Excel | Aspose.Cells DeleteRows method example | Update formulas after deleting rows in Excel .NET
// Developer Intent: Remove rows 20‑30 from a worksheet and have every dependent formula automatically corrected.
// Use Cases: Clean imported datasets by cutting out unwanted row blocks while preserving calculation integrity. | Prepare a report template: delete placeholder rows before inserting new data, ensuring formulas recalculate correctly. | Programmatic sheet restructuring where a specific row range must be removed and all related calculations stay accurate.
// AI Prompts: Write C# code using Aspose.Cells to delete rows 20‑30 and automatically adjust all formulas. | Explain the impact of each parameter in Cells.DeleteRows on formula references. | Show how to delete rows while preserving conditional formatting, named ranges, and chart data sources with Aspose.Cells.

using Aspose.Cells;

// Load a workbook, remove rows 20‑30 (zero‑based index 19, count 11) using Cells.DeleteRows, set the third parameter to true so all formula references adjust automatically, then save the file. Includes a ready‑to‑run Aspose.Cells .NET example.
class DeleteRowsExample
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet you need)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Delete rows 20 through 30 (1‑based indexing).
        // Zero‑based start index is 19 and the total rows to delete is 11.
        // The third parameter 'true' updates all formula references automatically.
        cells.DeleteRows(19, 11, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
