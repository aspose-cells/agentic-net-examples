// Title: C# – Merge Excel workbooks with Aspose.Cells and set Author to current Windows user
// Description: Shows how to load two .xlsx files, copy every worksheet from a source workbook into a target workbook, set the merged workbook's Author property to Environment.UserName, and save the result.
// Keywords: Aspose.Cells merge workbooks C# | set workbook author programmatically | Workbook.Settings.Author | Environment.UserName Excel metadata | combine Excel files Aspose.Cells | update document properties after merge
// Common Searches: how to set author of merged Excel file using Aspose.Cells | C# merge two workbooks and assign current user as author | Aspose.Cells copy worksheets and update metadata | set document properties after workbook merge .NET
// Developer Intent: Assign the current Windows user name to the Author property of a workbook after merging worksheets.
// Use Cases: Consolidate departmental reports while automatically recording the user who performed the merge for audit trails. | Run a nightly job that aggregates daily logs into a single workbook and tags the file with the service account name. | Create a reusable template that merges source workbooks and updates ownership information via the Author field.
// AI Prompts: Generate C# code that merges two Excel workbooks with Aspose.Cells and sets Workbook.Settings.Author to Environment.UserName. | Provide an example that copies all worksheets from a source workbook to a target workbook, then updates the Author, Title, and Company properties before saving. | Explain error handling for missing source files while still ensuring the Author property is set correctly after a merge.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    // Shows how to load two .xlsx files, copy every worksheet from a source workbook into a target workbook, set the merged workbook's Author property to Environment.UserName, and save the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Define file paths
                string targetPath = "Source1.xlsx";
                string sourcePath = "Source2.xlsx";
                string outputPath = "MergedWorkbook.xlsx";

                // Verify that input files exist
                if (!File.Exists(targetPath))
                    throw new FileNotFoundException($"Target workbook not found: {targetPath}");
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

                // Load the target workbook
                Workbook targetWorkbook = new Workbook(targetPath);

                // Load the source workbook to be merged
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Merge worksheets from source into target
                foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
                {
                    // Add a copy of each source worksheet to the target workbook using the sheet name
                    targetWorkbook.Worksheets.AddCopy(sourceSheet.Name);
                }

                // Set the author of the merged workbook
                targetWorkbook.Settings.Author = Environment.UserName;

                // Save the merged workbook
                targetWorkbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbooks merged successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
