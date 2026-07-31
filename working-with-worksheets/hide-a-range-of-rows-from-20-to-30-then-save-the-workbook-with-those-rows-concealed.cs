// Title: Hide rows 20‑30 in an Excel worksheet with Aspose.Cells for .NET and save the file
// Description: This C# example creates a new Workbook, accesses the first Worksheet, hides rows 20‑30 using Cells.HideRows (zero‑based start index 19, count 11), and saves the result as HiddenRows.xlsx.
// Keywords: Aspose.Cells hide rows | C# hide rows Excel | Cells.HideRows method | Excel row concealment .NET | save workbook with hidden rows | Aspose.Cells row visibility | hide multiple rows Aspose
// Common Searches: Aspose.Cells hide rows 20-30 | C# hide specific rows in Excel with Aspose | How to conceal rows in an Excel file using Aspose.Cells | Save Excel after hiding rows .NET | Cells.HideRows usage example
// Developer Intent: Hide a specific range of rows in an Excel worksheet and persist the hidden state when saving.
// Use Cases: Prepare a financial report where calculation rows are hidden before distribution. | Create a template that keeps internal notes hidden while showing data to users. | Generate print‑ready Excel files by hiding rows that should not appear on paper. | Automate data‑cleaning scripts that temporarily hide helper rows.
// AI Prompts: Provide C# code that uses Aspose.Cells to hide rows 20‑30 and save the workbook as HiddenRows.xlsx. | Explain the parameters of Cells.HideRows, including start index and row count, with examples. | Show how to hide rows by setting Row.IsHidden for rows 20‑30 as an alternative to Cells.HideRows. | Demonstrate how to unhide previously hidden rows using Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example creates a new Workbook, accesses the first Worksheet, hides rows 20‑30 using Cells.HideRows (zero‑based start index 19, count 11), and saves the result as HiddenRows.xlsx.
class HideRowsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 20 to 30 (zero‑based index starts at 19, total 11 rows)
        worksheet.Cells.HideRows(19, 11);

        // Save the workbook with the rows concealed
        workbook.Save("HiddenRows.xlsx");
    }
}
