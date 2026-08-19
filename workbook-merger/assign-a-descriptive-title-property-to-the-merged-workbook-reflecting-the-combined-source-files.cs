// Title: Assign a Descriptive Title to a Merged Workbook with Aspose.Cells for .NET
// Description: Shows how to merge multiple Excel files using Aspose.Cells Workbook.Combine, create a title that lists the source filenames, set the BuiltInDocumentProperties.Title, and save the combined workbook.
// Keywords: Aspose.Cells | C# merge workbooks | Workbook.Combine | set workbook title | BuiltInDocumentProperties | Excel consolidation | document properties .NET | combined workbook title
// Common Searches: Aspose.Cells set title after merging workbooks | C# combine Excel files and update document properties | How to add a title to a merged workbook using Aspose.Cells | Workbook.Combine title property | Set BuiltInDocumentProperties.Title in C#
// Developer Intent: Programmatically add a clear, source‑file list title to a workbook created by merging several Excel files.
// Use Cases: Consolidate departmental spreadsheets into a single report and embed the source file names in the document title for quick reference. | Generate a combined financial statement where the title records each input file, supporting audit trails and version control. | Automate monthly data aggregation, saving the merged workbook with a descriptive title that aids document management systems.
// AI Prompts: Write C# code using Aspose.Cells to merge an array of Excel files and set the workbook's BuiltInDocumentProperties.Title to a comma‑separated list of the source filenames. | Create a method that accepts a list of file paths, combines them into one workbook with Workbook.Combine, and updates the Title property with a custom prefix and the file names. | Explain how to retrieve, modify, and persist the Title property of a workbook after merging multiple workbooks with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to merge multiple Excel files using Aspose.Cells Workbook.Combine, create a title that lists the source filenames, set the BuiltInDocumentProperties.Title, and save the combined workbook.
class MergeWorkbooksWithTitle
{
    static void Main()
    {
        // Source workbook file paths to be merged
        string[] sourceFiles = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

        // Create the destination workbook using the first source file
        Workbook mergedWorkbook = new Workbook(sourceFiles[0]);

        // Combine the remaining workbooks into the destination workbook
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            Workbook wb = new Workbook(sourceFiles[i]);
            mergedWorkbook.Combine(wb);
        }

        // Build a descriptive title that lists the combined source files
        string descriptiveTitle = "Combined Workbook: " + string.Join(", ", sourceFiles);
        mergedWorkbook.BuiltInDocumentProperties.Title = descriptiveTitle;

        // Save the merged workbook to a new file
        string outputPath = "CombinedWorkbook.xlsx";
        mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

        // Optional: display confirmation
        Console.WriteLine("Workbook merged and saved to '" + outputPath + "'.");
        Console.WriteLine("Document Title set to: " + mergedWorkbook.BuiltInDocumentProperties.Title);
    }
}
