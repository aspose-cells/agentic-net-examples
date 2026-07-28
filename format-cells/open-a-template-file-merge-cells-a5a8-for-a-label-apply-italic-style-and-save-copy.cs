// Title: Merge A5:A8, Apply Italic Font, and Save a Copy with Aspose.Cells for .NET
// Description: Loads Template.xlsx, merges cells A5‑A8 on the first worksheet, sets the merged cell’s font to italic, and saves the result as Copy.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | italic font style Aspose.Cells | save workbook copy C# | merge range A5:A8 | format cells Aspose.Cells | C# Excel cell merging | Aspose.Cells style merged cells | Excel template modification Aspose
// Common Searches: Aspose.Cells merge cells A5 to A8 C# | how to set italic style on merged cells using Aspose.Cells | save modified Excel workbook as new file Aspose.Cells .NET | C# merge vertical range and apply font style Aspose | Aspose.Cells example merge and style cells
// Developer Intent: Merge cells A5‑A8, apply italic formatting, and export the workbook as a new file.
// Use Cases: Create a styled header in a generated report by merging rows and applying italic text. | Produce multiple customized copies of a template where each copy has an italic heading spanning several rows. | Design printable forms where the title occupies A5:A8 with italic formatting and the file is saved under a new name.
// AI Prompts: Write C# code using Aspose.Cells to merge cells A5:A8, set the font to italic, and save the workbook as a new file. | Show how to apply additional styles (e.g., font color, size) to a merged range and export the workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads Template.xlsx, merges cells A5‑A8 on the first worksheet, sets the merged cell’s font to italic, and saves the result as Copy.xlsx using Aspose.Cells for .NET.
class MergeAndStyleExample
{
    static void Main()
    {
        // Path to the existing template workbook
        string templatePath = "Template.xlsx";

        // Load the template workbook
        Workbook workbook = new Workbook(templatePath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells A5:A8 (zero‑based indices: row 4, column 0, 4 rows, 1 column)
        worksheet.Cells.Merge(firstRow: 4, firstColumn: 0, totalRows: 4, totalColumns: 1);

        // Apply italic font style to the merged cell (top‑left cell of the range)
        Style mergedStyle = worksheet.Cells[4, 0].GetStyle();
        mergedStyle.Font.IsItalic = true;
        worksheet.Cells[4, 0].SetStyle(mergedStyle);

        // Save the modified workbook as a new copy
        string outputPath = "Copy.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
