// Title: Automatically adjust row heights in every sheet of a merged Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a merged Excel file with Aspose.Cells, loops through all worksheets, calls AutoFitRows on each sheet, and saves the modified workbook. | Show how to apply Aspose.Cells' AutoFitRows method to every worksheet after combining multiple workbooks, ensuring row heights match cell content.
// Common Searches: Aspose.Cells C# auto fit rows after merging multiple workbooks | How to set row height automatically for all sheets in a merged Excel file using .NET | C# code example for applying AutoFitRows to each worksheet in a combined workbook | Adjust row heights in merged workbook with Aspose.Cells AutoFitRows | Iterate through worksheets to auto-fit rows in a merged Excel document using Aspose.Cells
// Tags: auto-fit rows Aspose.Cells .NET | merged workbook row height adjustment | apply AutoFitRows across all sheets | save workbook after row auto-fit | Aspose.Cells C# row height optimization

using Aspose.Cells;

// Loads a merged Excel workbook, auto-fits rows in every worksheet using Aspose.Cells, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the merged workbook
        Workbook workbook = new Workbook("MergedWorkbook.xlsx");

        // AutoFit all rows in each worksheet to adjust row height for content
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.AutoFitRows();
        }

        // Save the updated workbook
        workbook.Save("MergedWorkbook_AutoFitRows.xlsx");
    }
}
