// Title: C# – Merge H1:H4, apply light‑blue fill, and save as ODS using Aspose.Cells
// Description: Shows how to create a new workbook with Aspose.Cells for .NET, merge the range H1:H4, set a solid light‑blue background on the merged cell, and export the result as an ODS file via OdsSaveOptions (LibreOffice generator).
// Keywords: Aspose.Cells | C# | merge cells | cell background color | light blue fill | ODS export | OdsSaveOptions | LibreOffice generator | worksheet styling | Excel to ODS conversion
// Common Searches: Aspose.Cells merge cells C# example | Set background color for merged cells Aspose.Cells | Save workbook as ODS using Aspose.Cells .NET | Apply solid fill to a cell range in ODS with Aspose | C# code to create ODS file with styled header
// Developer Intent: Create a workbook, merge H1:H4, color it light blue, and save as ODS.
// Use Cases: Generate a branded ODS report where the title spans H1:H4 with a colored background for visual emphasis. | Automate production of spreadsheet templates that require a merged, colored header before distribution as ODS files. | Export data from a .NET application to ODS while preserving custom cell styling for downstream analysis.
// AI Prompts: Provide C# code to merge cells H1:H4, set a solid light‑blue background, and save the workbook as ODS with Aspose.Cells. | How do I use OdsSaveOptions to export a styled worksheet to ODS in Aspose.Cells for .NET? | Explain the steps to apply a background color to a merged cell range and generate an ODS file using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to create a new workbook with Aspose.Cells for .NET, merge the range H1:H4, set a solid light‑blue background on the merged cell, and export the result as an ODS file via OdsSaveOptions (LibreOffice generator).
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells H1:H4 (zero‑based indices: row 0, column 7, 4 rows, 1 column)
        cells.Merge(0, 7, 4, 1);

        // Set the background color of the merged cell to light blue
        Style style = cells[0, 7].GetStyle();
        style.ForegroundColor = Color.LightBlue;
        style.Pattern = BackgroundType.Solid;
        cells[0, 7].SetStyle(style);

        // Save the workbook as ODS using OdsSaveOptions
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice
        };
        workbook.Save("MergedLightBlue.ods", saveOptions);
    }
}
