// Title: C# – Hide rows 5‑15 and show formulas with Aspose.Cells before saving
// Description: Creates a new workbook, hides rows 5‑15 on the first worksheet using zero‑based indexing, switches the sheet to formula view, and saves the file as HiddenRowsAndFormulas.xlsx.
// Keywords: Aspose.Cells hide rows C# | show formulas Aspose.Cells | HideRows method | ShowFormulas property | .NET Excel hide rows | display formulas before save | Aspose.Cells worksheet formatting
// Common Searches: Aspose.Cells hide rows 5 to 15 C# | How to enable formula view in Aspose.Cells | Hide multiple rows and show formulas Aspose.Cells .NET | C# code to hide rows and display formulas in Excel workbook | Aspose.Cells hide rows and show formulas example
// Developer Intent: Hide rows 5‑15 and enable formula view on the same worksheet before saving the workbook.
// Use Cases: Financial reports where intermediate calculation rows are hidden but formulas stay visible for auditors. | Template generation that conceals data rows while exposing underlying formulas for user transparency. | Debug builds that hide helper rows yet display formulas to simplify troubleshooting.
// AI Prompts: Generate C# code using Aspose.Cells to hide rows 5‑15, turn on formula view, and save the workbook. | Show how to use HideRows and ShowFormulas together in Aspose.Cells for .NET. | Explain the steps to conceal a range of rows and display formulas on a single worksheet without affecting others.

using System;
using Aspose.Cells;

// Creates a new workbook, hides rows 5‑15 on the first worksheet using zero‑based indexing, switches the sheet to formula view, and saves the file as HiddenRowsAndFormulas.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 5 through 15 (zero‑based index: start at 4, hide 11 rows)
        worksheet.Cells.HideRows(4, 11);

        // Enable formula view on the worksheet
        worksheet.ShowFormulas = true;

        // Save the workbook
        workbook.Save("HiddenRowsAndFormulas.xlsx");
    }
}
