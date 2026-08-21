// Title: Split and Freeze Panes to Create a Fixed Header in Excel with Aspose.Cells for .NET (C#)
// Description: Shows how to split a worksheet window and then freeze the top row as a fixed header using Aspose.Cells for .NET. The example fills 30 rows, calls Worksheet.Split(), applies Worksheet.FreezePanes(1,0,1,0), and saves the workbook as SplitAndFreezeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | split panes | freeze panes | fixed header | Worksheet.Split | Worksheet.FreezePanes | Excel automation | Excel UI simulation
// Common Searches: Aspose.Cells split panes then freeze header row C# | How to freeze top row after splitting worksheet with Aspose.Cells | Split and freeze panes on same sheet Aspose.Cells .NET | Create fixed header area in Excel using Aspose.Cells | Worksheet.Split and FreezePanes example C#
// Developer Intent: The developer wants to divide the worksheet view into panes and then lock the first row so it remains visible while scrolling.
// Use Cases: Building a reporting workbook where column titles stay in view during vertical scrolling. | Designing a large data grid that allows side‑by‑side comparison after splitting the view, with a frozen header for context. | Preparing an interactive Excel file for data analysis where the header row is locked while users navigate through rows.
// AI Prompts: Provide C# code that splits a worksheet at a specific row and then freezes the top row using Aspose.Cells for .NET. | Show an example of combining Worksheet.Split() and Worksheet.FreezePanes() on the same sheet with custom frozen rows and columns. | Explain the difference between Split() and FreezePanes() in Aspose.Cells and how to use them together to create a fixed header area.

using Aspose.Cells;

// Shows how to split a worksheet window and then freeze the top row as a fixed header using Aspose.Cells for .NET. The example fills 30 rows, calls Worksheet.Split(), applies Worksheet.FreezePanes(1,0,1,0), and saves the workbook as SplitAndFreezeDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to illustrate the header row
        for (int i = 0; i < 30; i++)
        {
            sheet.Cells[i, 0].PutValue($"Item {i + 1}");
            sheet.Cells[i, 1].PutValue(i * 5);
        }

        // Split the window to create multiple panes
        sheet.Split();

        // Freeze the top row (row index 1) so the header stays visible while scrolling
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        sheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook
        workbook.Save("SplitAndFreezeDemo.xlsx");
    }
}
