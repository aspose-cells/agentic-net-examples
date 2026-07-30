// Title: C# – Apply Aspose.Cells AutoFilter to show only rows where Status = "Completed"
// Description: Creates a workbook, populates it with task data, sets an AutoFilter on range A1:D5, filters the Status column (index 2) for the value "Completed", refreshes the view to hide non‑matching rows, optionally checks hidden rows, and saves the file as FilteredByStatusCompleted.xlsx.
// Keywords: Aspose.Cells C# | AutoFilter Excel .NET | filter rows by column value | hide rows programmatically | Excel status completed filter | worksheet filtering Aspose.Cells | C# Excel automation | filter Excel sheet by criteria | Aspose.Cells hide non‑matching rows | apply AutoFilter with Aspose
// Common Searches: Aspose.Cells filter rows where Status = Completed | C# hide rows that do not meet AutoFilter criteria | How to use AutoFilter in Aspose.Cells .NET | Programmatically show only completed tasks in Excel using Aspose | Apply Excel AutoFilter with Aspose.Cells C# example
// Developer Intent: Apply an AutoFilter so that only rows with Status equal to "Completed" stay visible, while all other rows are hidden.
// Use Cases: Generate a task report that includes only completed items. | Create a clean dashboard view by automatically hiding pending or in‑progress rows. | Export a filtered list of completed entries for downstream processing or analytics.
// AI Prompts: Show how to add a second filter condition (e.g., Status = "Completed" OR Owner = "Alice"). | Provide code to clear the AutoFilter and reveal all rows again. | Explain how to retrieve the row indices that remain visible after filtering.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates it with task data, sets an AutoFilter on range A1:D5, filters the Status column (index 2) for the value "Completed", refreshes the view to hide non‑matching rows, optionally checks hidden rows, and saves the file as FilteredByStatusCompleted.xlsx.
    public class FilterStatusCompletedDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Task");
            cells["C1"].PutValue("Status");   // Status column (index 2)
            cells["D1"].PutValue("Owner");

            // Data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Design");
            cells["C2"].PutValue("Completed");
            cells["D2"].PutValue("Alice");

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Development");
            cells["C3"].PutValue("In Progress");
            cells["D3"].PutValue("Bob");

            cells["A4"].PutValue(3);
            cells["B4"].PutValue("Testing");
            cells["C4"].PutValue("Completed");
            cells["D4"].PutValue("Charlie");

            cells["A5"].PutValue(4);
            cells["B5"].PutValue("Deployment");
            cells["C5"].PutValue("Pending");
            cells["D5"].PutValue("Diana");

            // Apply AutoFilter to the range that includes the header and all data rows
            sheet.AutoFilter.Range = "A1:D5";

            // Filter the Status column (field index 2) for the value "Completed"
            sheet.AutoFilter.Filter(2, "Completed");

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Verify which rows are hidden (optional)
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                bool hidden = sheet.Cells.IsRowHidden(row);
                Console.WriteLine($"Row {row + 1} hidden: {hidden}");
            }

            // Save the workbook
            string outputPath = "FilteredByStatusCompleted.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
