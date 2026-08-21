// Title: Merge AB6:AC8 and apply phone number format (###) ###‑#### with Aspose.Cells for .NET
// Description: Loads a workbook, merges the range AB6:AC8, applies the custom number format "(###) ###‑####" to the merged cell, and saves the result. Demonstrates zero‑based indexing, style handling, and export in C#.
// Keywords: Aspose.Cells merge cells C# | custom phone number format | set number format merged cell | Aspose.Cells export workbook | C# Excel cell merging
// Common Searches: Aspose.Cells merge range AB6 AC8 C# | apply custom phone number format with Aspose.Cells | how to set number format on merged cells Aspose.Cells .NET | save workbook after merging cells Aspose.Cells
// Developer Intent: Combine a specific cell block, format it as a phone number, and write the workbook to disk.
// Use Cases: Create a contact sheet where each phone number spans two columns with a standardized format. | Generate a report header that merges cells for phone numbers and enforces consistent formatting before distribution. | Design an invoice template with a merged phone‑number field that automatically displays numbers as (###) ###‑####.
// AI Prompts: Show C# code to merge AB6:AC8 and set the phone number format (###) ###‑#### using Aspose.Cells. | Explain how to apply a custom number format to a merged cell in Aspose.Cells for .NET. | What are the zero‑based row and column indices for merging cells AB6:AC8 in Aspose.Cells?

using System;
using Aspose.Cells;

// Loads a workbook, merges the range AB6:AC8, applies the custom number format "(###) ###‑####" to the merged cell, and saves the result. Demonstrates zero‑based indexing, style handling, and export in C#.
class MergeAndFormatPhoneNumber
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells AB6:AC8
        // AB -> column index 27, AC -> column index 28
        // Row 6 -> row index 5 (zero‑based)
        // Total rows = 3 (6,7,8), total columns = 2 (AB,AC)
        worksheet.Cells.Merge(5, 27, 3, 2);

        // Apply phone number format "(###) ###‑####" to the merged cell (top‑left cell of the range)
        Style phoneStyle = worksheet.Cells[5, 27].GetStyle();
        phoneStyle.Custom = "(###) ###-####";
        worksheet.Cells[5, 27].SetStyle(phoneStyle);

        // Save the modified workbook (replace with desired output path)
        workbook.Save("OutputWorkbook.xlsx");
    }
}
