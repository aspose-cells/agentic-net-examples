// Title: C# – Hide Rows 5‑15 and Show Formulas in an Aspose.Cells Worksheet
// Description: Creates a new workbook, inserts sample data and a formula, hides rows 5‑15 with Cells.HideRows, enables worksheet.ShowFormulas to display formulas instead of results, and saves the file as HiddenRows_ShowFormulas.xlsx.
// Keywords: Aspose.Cells hide rows C# | ShowFormulas property | display formulas Aspose.Cells | hide specific rows .NET | Excel row visibility | Aspose.Cells workbook save
// Common Searches: Aspose.Cells hide rows 5 to 15 C# | How to enable formula view in Aspose.Cells | Hide multiple rows and show formulas in .NET Excel | Aspose.Cells hide rows and display formulas before saving
// Developer Intent: Hide rows 5‑15 and turn on formula view before saving the workbook.
// Use Cases: Prepare a financial report where input rows are concealed but calculation formulas stay visible for auditors. | Generate a template that hides raw data rows while exposing the underlying formulas to end‑users. | Distribute an Excel file with selected rows hidden and formulas displayed to simplify review and reduce accidental edits.
// AI Prompts: Provide C# code using Aspose.Cells to hide rows 5‑15 and set ShowFormulas = true before saving. | Show how to conceal a range of rows and toggle formula display in an Aspose.Cells worksheet. | Explain the steps to hide specific rows and enable formula view without altering existing cell values in a .NET workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsHideRowsAndShowFormulas
{
    // Creates a new workbook, inserts sample data and a formula, hides rows 5‑15 with Cells.HideRows, enables worksheet.ShowFormulas to display formulas instead of results, and saves the file as HiddenRows_ShowFormulas.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example data with a formula (optional, just to illustrate ShowFormulas)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=A1+A2"; // Formula cell

            // Hide rows 5 through 15 (zero‑based index: start at 4, hide 11 rows)
            cells.HideRows(4, 11);

            // Enable formula view on the worksheet
            worksheet.ShowFormulas = true;

            // Save the workbook
            workbook.Save("HiddenRows_ShowFormulas.xlsx");
        }
    }
}
