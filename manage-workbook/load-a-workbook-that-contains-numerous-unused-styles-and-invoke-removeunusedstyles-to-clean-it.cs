// Title: How to remove unused cell styles from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that accepts input and output .xlsx file paths, loads the workbook with Aspose.Cells, calls RemoveUnusedStyles, and saves the cleaned file. | Show a .NET console example that iterates over a folder of Excel files and removes unused styles from each workbook using Aspose.Cells. | Demonstrate adding Workbook.RemoveUnusedStyles() into an existing Aspose.Cells workflow that also modifies worksheets before saving.
// Common Searches: C# Aspose.Cells remove unused cell styles from an existing .xlsx file | How to clean up Excel workbook formatting with Aspose.Cells RemoveUnusedStyles method | Reduce .xlsx file size by deleting unused styles using Aspose.Cells in .NET | Batch remove unused styles from multiple Excel workbooks with Aspose.Cells C# | Aspose.Cells RemoveUnusedStyles example code for .NET developers
// Tags: Aspose.Cells RemoveUnusedStyles method | C# delete unused cell styles | optimize .xlsx file size Aspose.Cells | clean Excel workbook formatting .NET | batch process Excel files Aspose.Cells

using Aspose.Cells;

// Loads 'input.xlsx' with Aspose.Cells, removes all unused cell styles via Workbook.RemoveUnusedStyles(), and saves the cleaned workbook as 'output.xlsx' using C#.
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
