// Title: Merge P2:P5, Enable Text Wrap, Apply AutoFilter and Save as XLSX with Aspose.Cells (C#)
// Description: Shows how to use Aspose.Cells for .NET to create a workbook, merge the range P2:P5, set the IsTextWrapped style, add sample text, apply an AutoFilter to the same column, and save the file in XLSX format.
// Keywords: Aspose.Cells | C# | .NET | merge cells P2:P5 | wrap text merged cells | auto filter Aspose.Cells | save workbook as XLSX | Excel automation | worksheet style | Excel filter programmatically
// Common Searches: Aspose.Cells merge cells and wrap text C# | How to add AutoFilter to a merged range with Aspose.Cells | Save merged and wrapped cells as XLSX using Aspose.Cells .NET | C# code for merging P2:P5 and applying filter in Excel
// Developer Intent: Create an XLSX file where cells P2 through P5 are merged, the contained text wraps automatically, and an AutoFilter is active on that column.
// Use Cases: Design a report header that spans multiple rows with wrapped description text while keeping the column filterable. | Build an export template where a merged title cell needs line breaks and the data column must support filtering. | Generate a printable worksheet with a section title in merged cells and enable downstream data analysis via AutoFilter.
// AI Prompts: Provide C# Aspose.Cells code to merge cells P2:P5, enable text wrap, add an AutoFilter, and save as XLSX. | Explain how to apply the IsTextWrapped style to a merged range and refresh the AutoFilter in Aspose.Cells for .NET. | Step‑by‑step guide for creating a workbook with merged, wrapped cells and an active filter using Aspose.Cells in C#.

using Aspose.Cells;

// Shows how to use Aspose.Cells for .NET to create a workbook, merge the range P2:P5, set the IsTextWrapped style, add sample text, apply an AutoFilter to the same column, and save the file in XLSX format.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells P2:P5
        // Row indices are zero‑based: P2 is row 1, column 15
        // Merge 4 rows (P2 to P5) and 1 column
        worksheet.Cells.Merge(1, 15, 4, 1);

        // Set wrap text on the merged cell (top‑left cell of the range)
        Style style = worksheet.Cells[1, 15].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells[1, 15].SetStyle(style);

        // Add some sample text to demonstrate wrapping
        worksheet.Cells[1, 15].PutValue("This is a long text that will be wrapped inside the merged cells P2:P5.");

        // Apply an auto‑filter to the same column range
        worksheet.AutoFilter.Range = "P2:P5";
        worksheet.AutoFilter.Refresh();

        // Save the workbook as XLSX
        workbook.Save("MergedWrappedAutoFilter.xlsx", SaveFormat.Xlsx);
    }
}
