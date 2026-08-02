// Title: Aspose.Cells .NET: Merge H1:H4, set light‑blue fill, export to ODS
// Description: C# example that creates a new Workbook, merges cells H1 through H4 on the first worksheet, applies a solid light‑blue background to the merged range, and saves the result as an ODS file with OdsSaveOptions.
// Keywords: Aspose.Cells C# | merge cells | cell background color | light blue fill | ODS export .NET | OdsSaveOptions | style merged cells | Aspose.Cells example
// Common Searches: Aspose.Cells merge cells H1:H4 | Set background color for merged cells Aspose.Cells C# | Save workbook as ODS using Aspose.Cells .NET | Apply solid fill to merged range Aspose.Cells | C# ODS export with cell styling Aspose
// Developer Intent: Create a workbook, merge H1:H4, color the merged cell light blue, and export it as an ODS document using Aspose.Cells for .NET.
// Use Cases: Generate a multi‑row header with a light‑blue background for reports that need to be opened in LibreOffice. | Build a reusable template where the title cell spans H1:H4, is styled with a blue fill, and is saved as ODS for cross‑platform distribution.
// AI Prompts: Write C# code with Aspose.Cells to merge cells A1:C1, apply a yellow solid fill, and save the workbook as an ODS file. | Explain how to set a custom background color on a merged cell range and export the worksheet to ODS using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// C# example that creates a new Workbook, merges cells H1 through H4 on the first worksheet, applies a solid light‑blue background to the merged range, and saves the result as an ODS file with OdsSaveOptions.
class MergeAndStyleExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells H1:H4 (zero‑based indices: row 0, column 7, 4 rows, 1 column)
        worksheet.Cells.Merge(0, 7, 4, 1);

        // Apply light blue background to the merged cell
        Style style = worksheet.Cells[0, 7].GetStyle();
        style.ForegroundColor = Color.LightBlue;
        style.Pattern = BackgroundType.Solid;
        worksheet.Cells[0, 7].SetStyle(style);

        // Save the workbook as ODS using OdsSaveOptions
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        workbook.Save("MergedCellsLightBlue.ods", saveOptions);
    }
}
