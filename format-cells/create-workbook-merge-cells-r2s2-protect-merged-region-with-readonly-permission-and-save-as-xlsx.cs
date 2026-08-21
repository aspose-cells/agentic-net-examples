// Title: Aspose.Cells for .NET – Merge R2:S2, set read‑only protection, and export to XLSX
// Description: Shows how to create a new Workbook, merge the range R2:S2 on the first worksheet, apply worksheet protection so the merged cells become read‑only, and save the result as MergedReadOnly.xlsx in XLSX format using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | R2:S2 | read‑only protection | worksheet protect | C# Excel export | save as XLSX | Aspose.Cells .NET example | protect merged region | Excel cell lock | Aspose.Cells API
// Common Searches: Aspose.Cells merge cells R2:S2 C# | How to protect merged cells in Aspose.Cells | Make a merged range read‑only with Aspose.Cells .NET | Save protected Excel workbook using Aspose.Cells | C# code to lock specific cells in Excel
// Developer Intent: Create a workbook, merge R2:S2, protect that range as read‑only, and save the file as XLSX.
// Use Cases: Design a template where the header spanning R2:S2 is locked to prevent user edits. | Generate a report with a merged title row that remains immutable while other cells stay editable. | Build a data‑entry sheet where the merged title row is protected to preserve formatting and wording.
// AI Prompts: Write C# code with Aspose.Cells that merges R2:S2, applies read‑only protection only to that merged area, and saves the workbook as an XLSX file. | Explain how worksheet protection works in Aspose.Cells and how to allow editing of all cells except a specific merged region. | Provide an example of using ProtectionType options after merging cells in Aspose.Cells for .NET to lock the merged range while keeping the rest of the sheet editable.

using Aspose.Cells;

// Shows how to create a new Workbook, merge the range R2:S2 on the first worksheet, apply worksheet protection so the merged cells become read‑only, and save the result as MergedReadOnly.xlsx in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells R2:S2 (zero‑based indices: row 1, column 17, 1 row, 2 columns)
        sheet.Cells.Merge(1, 17, 1, 2);

        // Protect the worksheet so that the merged region becomes read‑only
        sheet.Protect(ProtectionType.All);

        // Save the workbook as XLSX
        workbook.Save("MergedReadOnly.xlsx", SaveFormat.Xlsx);
    }
}
