// Title: C# – Merge Excel workbooks and AutoFit all columns using Aspose.Cells
// Description: Loads two Excel files, combines the second workbook into the first with Workbook.Combine, auto‑fits every column in each worksheet via Worksheet.AutoFitColumns, and saves the merged workbook. Includes file‑existence checks and error handling.
// Keywords: Aspose.Cells merge workbooks | C# combine Excel files | AutoFitColumns after combine | adjust column width Aspose.Cells | .NET Excel workbook merge | auto fit columns C# | Workbook.Combine example
// Common Searches: how to merge two Excel workbooks in C# with Aspose.Cells | auto fit columns after combining workbooks Aspose.Cells .NET | C# code to combine Excel files and adjust column widths automatically | Aspose.Cells Workbook.Combine with AutoFitColumns | merge multiple Excel sheets and auto‑size columns using Aspose
// Developer Intent: Combine several Excel workbooks into one and automatically size all columns for optimal readability.
// Use Cases: Consolidate monthly sales reports into a single workbook with columns sized to content. | Create a master data file from multiple source templates, ensuring each sheet is properly formatted. | Automate generation of a combined financial statement where column widths adapt to the merged data.
// AI Prompts: Generate C# code that merges an arbitrary number of Excel workbooks with Aspose.Cells and applies AutoFitColumns to every worksheet. | Explain performance considerations when combining large workbooks and auto‑fitting columns in Aspose.Cells. | Show how to log the original and new column widths after merging workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads two Excel files, combines the second workbook into the first with Workbook.Combine, auto‑fits every column in each worksheet via Worksheet.AutoFitColumns, and saves the merged workbook. Includes file‑existence checks and error handling.
class Program
{
    static void Main()
    {
        const string firstPath = "InputWorkbook1.xlsx";
        const string secondPath = "InputWorkbook2.xlsx";
        const string outputPath = "MergedWorkbook_AutoFitted.xlsx";

        // Verify that the input files exist before attempting to load them
        if (!File.Exists(firstPath) || !File.Exists(secondPath))
        {
            Console.WriteLine("One or both input workbook files were not found.");
            Console.WriteLine($"Expected: {firstPath} and {secondPath}");
            return;
        }

        try
        {
            // Load the first workbook (source)
            Workbook mergedWorkbook = new Workbook(firstPath);

            // Load the second workbook to be merged
            Workbook secondWorkbook = new Workbook(secondPath);

            // Combine the second workbook into the first one
            mergedWorkbook.Combine(secondWorkbook);

            // AutoFit all columns in every worksheet of the merged workbook
            foreach (Worksheet sheet in mergedWorkbook.Worksheets)
            {
                sheet.AutoFitColumns(); // Adjust column widths based on content
            }

            // Save the resulting merged workbook
            mergedWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
