// Title: C# – Hide "Inactive" rows after CSV import using Aspose.Cells AutoFilter
// Description: Load a CSV into an Aspose.Cells workbook, define an AutoFilter range, apply a NotEqual filter on the Status column to hide rows with "Inactive", refresh the view, and save the result as XLSX.
// Keywords: Aspose.Cells AutoFilter C# | filter rows by value Aspose.Cells | hide inactive rows CSV | custom filter NotEqual Aspose.Cells | CSV to XLSX Aspose.Cells
// Common Searches: Aspose.Cells hide rows where status = Inactive | C# AutoFilter CSV Aspose.Cells example | filter out inactive records using Aspose.Cells | apply custom NotEqual filter Aspose.Cells .NET
// Developer Intent: Exclude rows with a status of "Inactive" after loading a CSV file.
// Use Cases: Clean a CSV export by removing inactive records before reporting. | Create a filtered Excel view for dashboards without modifying the source file. | Prepare data for downstream processing where only active entries are required.
// AI Prompts: Generate C# code that reads a CSV with Aspose.Cells, sets an AutoFilter to hide rows where column C equals "Inactive", and saves the workbook as XLSX. | Show how to use Aspose.Cells AutoFilter.Custom with FilterOperatorType.NotEqual on the third column to filter out "Inactive" rows.

using System;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    // Load a CSV into an Aspose.Cells workbook, define an AutoFilter range, apply a NotEqual filter on the Status column to hide rows with "Inactive", refresh the view, and save the result as XLSX.
    class Program
    {
        static void Main()
        {
            // Load CSV data into a workbook
            // (Assumes "data.csv" is in the same folder as the executable)
            Workbook workbook = new Workbook("data.csv");

            // Access the first worksheet where the CSV data is loaded
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
            int maxColumn = worksheet.Cells.MaxDataColumn; // zero‑based index of the last column with data

            // Set the AutoFilter range to include the header row (row 0) and all data rows
            // The range is defined by startRow, startColumn, endRow (the column is inferred from maxColumn)
            worksheet.AutoFilter.SetRange(0, 0, maxRow);

            // Assume the "Status" column is the third column (index 2, zero‑based)
            int statusColumnIndex = 2;

            // Apply a custom filter to hide rows where the status equals "Inactive"
            // Using the NotEqual operator will keep rows whose status is not "Inactive"
            worksheet.AutoFilter.Custom(statusColumnIndex, FilterOperatorType.NotEqual, "Inactive");

            // Refresh the filter to apply the changes (rows not matching the criteria become hidden)
            worksheet.AutoFilter.Refresh();

            // Save the filtered workbook to an XLSX file
            workbook.Save("FilteredOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}
