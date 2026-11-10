// Title: Remove unused cell styles from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: How to delete unused cell styles in an Excel file using Aspose.Cells C# | Aspose.Cells RemoveUnusedStyles method example for .NET | Programmatically clean up redundant styles in a .xlsx workbook with C# | Reduce Excel file size by removing unused styles using Aspose.Cells | Sample code to remove unused styles from a workbook in Aspose.Cells for .NET
// Tags: Aspose.Cells RemoveUnusedStyles API | unused Excel styles removal C# | optimize workbook size by cleaning styles .NET | delete redundant cell styles with Aspose.Cells | Excel style cleanup using Aspose.Cells

using Aspose.Cells;

// Loads an Excel file into an Aspose.Cells Workbook, calls RemoveUnusedStyles to purge all styles that are not referenced, and saves the cleaned workbook, helping to shrink file size and simplify style management.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all styles that are not used in the workbook
        workbook.RemoveUnusedStyles();

        // Save the cleaned workbook to a new file
        workbook.Save("output.xlsx");
    }
}
