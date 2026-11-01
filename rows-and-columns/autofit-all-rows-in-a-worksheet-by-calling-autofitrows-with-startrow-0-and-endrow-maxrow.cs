// Title: How to auto‑fit every row in an Aspose.Cells worksheet from the first row to the last populated row using C#
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, retrieves the maximum data row via MaxDataRow, and calls Worksheet.AutoFitRows(0, maxRow) to auto‑size all rows. | Create a C# example that checks if a worksheet contains data and, when it does, auto‑fits rows from index 0 through the last used row using Aspose.Cells.
// Common Searches: C# Aspose.Cells auto fit rows from first to last used row | How to use MaxDataRow with AutoFitRows in Aspose.Cells | Auto‑size all rows in an Excel sheet using Aspose.Cells .NET | Worksheet.AutoFitRows range example C# | Determine last populated row in Aspose.Cells worksheet
// Tags: Aspose.Cells row height auto‑adjustment | Worksheet.AutoFitRows using MaxDataRow | C# find last populated row Aspose.Cells | auto‑size all rows Excel .NET | Excel worksheet row auto‑fit C#

using Aspose.Cells;

// The program creates or loads a workbook, accesses the first worksheet, obtains the last row containing data with MaxDataRow, auto‑fits rows from index 0 to that row using Worksheet.AutoFitRows, and saves the result as AutoFitRowsResult.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create rule

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Find the last row that contains data
        int maxRow = sheet.Cells.MaxDataRow; // returns -1 if the sheet is empty

        // Auto‑fit all rows from the first row (0) to the last used row
        if (maxRow >= 0)
        {
            sheet.AutoFitRows(0, maxRow);
        }

        // Save the workbook
        workbook.Save("AutoFitRowsResult.xlsx"); // save rule
    }
}
