// Title: C# Aspose.Cells: AutoFilter to Show Only Rows with Status = 'Completed' and Hide the Rest
// Description: Creates a workbook, adds ID, Status, and Value columns, sets an AutoFilter on range A1:C5, filters the Status column for the value "Completed", refreshes the view to hide non‑matching rows, prints hidden‑row flags for verification, and saves the result as FilteredStatusCompleted.xlsx.
// Keywords: Aspose.Cells C# filter rows by column value | AutoFilter hide non‑matching rows Aspose.Cells | filter Excel rows status Completed .NET | programmatic Excel row hiding Aspose.Cells | C# Excel AutoFilter example
// Common Searches: Aspose.Cells filter rows where Status = Completed | C# hide rows that do not meet AutoFilter criteria | How to apply AutoFilter to a specific column in Aspose.Cells | Show only completed items in Excel using Aspose.Cells C# | Retrieve hidden rows after applying AutoFilter Aspose.Cells
// Developer Intent: Apply an AutoFilter that displays only rows whose Status column equals "Completed" and automatically hides every other row in a worksheet.
// Use Cases: Generate a report that lists only completed tasks before exporting to Excel. | Automate data validation by confirming which rows are hidden after a status filter. | Create a clean view for end‑users that excludes pending or in‑progress items.
// AI Prompts: Show how to extend the filter to include multiple statuses such as "Completed" and "InProgress" with Aspose.Cells. | Provide code to clear the AutoFilter and reveal all rows in the worksheet. | Explain how to obtain a collection of hidden row indices after applying an AutoFilter in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds ID, Status, and Value columns, sets an AutoFilter on range A1:C5, filters the Status column for the value "Completed", refreshes the view to hide non‑matching rows, prints hidden‑row flags for verification, and saves the result as FilteredStatusCompleted.xlsx.
    public class FilterStatusCompletedDemo
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a "Status" column (column B)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Status");
            sheet.Cells["C1"].PutValue("Value");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Completed");
            sheet.Cells["C2"].PutValue(100);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Pending");
            sheet.Cells["C3"].PutValue(200);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Completed");
            sheet.Cells["C4"].PutValue(150);

            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("InProgress");
            sheet.Cells["C5"].PutValue(120);

            // Define the autofilter range (including header row)
            sheet.AutoFilter.Range = "A1:C5";

            // Apply filter on the "Status" column (field index 1) for the value "Completed"
            sheet.AutoFilter.Filter(1, "Completed");

            // Refresh the filter to hide rows that do not match the criteria
            sheet.AutoFilter.Refresh();

            // Output hidden row information for verification
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                bool isHidden = sheet.Cells.IsRowHidden(row);
                Console.WriteLine($"Row {row + 1} hidden: {isHidden}");
            }

            // Save the filtered workbook
            workbook.Save("FilteredStatusCompleted.xlsx", SaveFormat.Xlsx);
        }
    }
}
