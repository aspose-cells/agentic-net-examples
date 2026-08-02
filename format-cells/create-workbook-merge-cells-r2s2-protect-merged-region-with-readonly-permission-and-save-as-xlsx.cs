// Title: Aspose.Cells for .NET – Merge cells R2:S2, lock the range read‑only, and export to XLSX
// Description: Learn how to create a new workbook with Aspose.Cells, merge the range R2:S2, protect that merged region as read‑only while keeping other cells editable, and save the result as an XLSX file using C#.
// Keywords: Aspose.Cells merge cells C# | protect merged range read only | R2:S2 merge Aspose.Cells | worksheet protection Aspose.Cells .NET | save workbook as XLSX C# | lock merged cells Excel | Aspose.Cells example read‑only merged cells
// Common Searches: how to merge R2:S2 with Aspose.Cells | make merged cells read‑only in Excel using C# | protect only a merged range Aspose.Cells .NET | save merged and locked cells as XLSX | Aspose.Cells worksheet protection example
// Developer Intent: Create a workbook, merge cells R2:S2, protect that merged area as read‑only, and save the file in XLSX format.
// Use Cases: Design a template where the title header (R2:S2) is merged and locked to prevent edits. | Distribute a report with a merged title cell that must remain unchanged after delivery. | Build a data‑entry form where instructional text in a merged cell is read‑only while the rest of the sheet stays editable.
// AI Prompts: Generate C# code with Aspose.Cells to merge cells R2:S2, lock the merged region, and save as an XLSX workbook. | Explain how worksheet protection works on merged cells in Aspose.Cells and how to allow editing of other cells. | Show an Aspose.Cells example that protects only the merged range R2:S2 while leaving the rest of the worksheet unprotected.

using Aspose.Cells;

// Learn how to create a new workbook with Aspose.Cells, merge the range R2:S2, protect that merged region as read‑only while keeping other cells editable, and save the result as an XLSX file using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells R2:S2 (row index 1, column index 17, spanning 1 row and 2 columns)
        sheet.Cells.Merge(1, 17, 1, 2);

        // Optional: put a value in the merged cell
        sheet.Cells[1, 17].PutValue("Read‑Only Merged");

        // Protect the worksheet so the merged region becomes read‑only
        sheet.Protect(ProtectionType.All);

        // Save the workbook as XLSX
        workbook.Save("MergedReadOnly.xlsx", SaveFormat.Xlsx);
    }
}
