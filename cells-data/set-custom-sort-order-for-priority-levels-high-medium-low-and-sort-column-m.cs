// Title: Custom priority sort (High, Medium, Low) for column M with Aspose.Cells DataSorter in C#
// Description: Creates a workbook, populates column M with priority values, enables header detection, adds a custom sort key ordered High,Medium,Low, defines the sort range, executes the sort, and saves the file as CustomPrioritySorted.xlsx.
// Keywords: Aspose.Cells | DataSorter | custom sort list | C# Excel sorting | priority order High Medium Low | .NET Excel automation | sort column with header | Excel custom list sort | Aspose.Cells example | C# workbook sorting
// Common Searches: Aspose.Cells custom sort order C# | How to sort Excel column by custom list using DataSorter | C# sort column M High Medium Low Aspose | Define custom sort keys in Aspose.Cells .NET | Excel priority sorting with Aspose.Cells
// Developer Intent: Implement a custom sort that ranks the values High, Medium, and Low in column M of an Excel worksheet using Aspose.Cells.
// Use Cases: Prioritize task lists before exporting to Excel. | Arrange support tickets by severity in a generated report. | Order product backlog items by custom priority levels for stakeholder review.
// AI Prompts: Generate C# code with Aspose.Cells to sort column D using a custom order Urgent,Normal,Low. | Show how to sort multiple columns, each with its own custom list, using DataSorter. | Add error handling for unexpected priority values when applying a custom sort list.

using System;
using Aspose.Cells;

// Creates a workbook, populates column M with priority values, enables header detection, adds a custom sort key ordered High,Medium,Low, defines the sort range, executes the sort, and saves the file as CustomPrioritySorted.xlsx.
class CustomPrioritySort
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column M (index 12)
        // Row 0 – header
        sheet.Cells[0, 12].PutValue("Priority");
        // Data rows
        sheet.Cells[1, 12].PutValue("Low");
        sheet.Cells[2, 12].PutValue("High");
        sheet.Cells[3, 12].PutValue("Medium");
        sheet.Cells[4, 12].PutValue("High");

        // Get the DataSorter object
        DataSorter sorter = workbook.DataSorter;

        // Indicate that the range contains a header row
        sorter.HasHeaders = true;

        // Define custom sort order: High > Medium > Low
        // Use the overload that accepts a custom list string
        sorter.AddKey(12, SortOrder.Ascending, "High,Medium,Low");

        // Define the area to sort (including header)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 12,
            EndRow = 4,
            EndColumn = 12
        };

        // Perform the sort
        sorter.Sort(sheet.Cells, sortArea);

        // Save the workbook
        workbook.Save("CustomPrioritySorted.xlsx");
    }
}
