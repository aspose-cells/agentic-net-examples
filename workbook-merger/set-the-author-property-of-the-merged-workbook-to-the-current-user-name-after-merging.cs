using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndSetAuthor
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string primaryPath = "PrimaryWorkbook.xlsx";

                // Verify primary workbook exists
                if (!File.Exists(primaryPath))
                    throw new FileNotFoundException($"Primary workbook not found: {primaryPath}");

                // Load the primary workbook (the one that will receive the merged sheets)
                Workbook mergedWorkbook = new Workbook(primaryPath);

                // Files to merge
                string[] filesToMerge = { "Workbook1.xlsx", "Workbook2.xlsx" };

                foreach (string file in filesToMerge)
                {
                    // Skip missing files
                    if (!File.Exists(file))
                    {
                        Console.WriteLine($"Warning: File not found and will be skipped: {file}");
                        continue;
                    }

                    // Load source workbook
                    Workbook sourceWorkbook = new Workbook(file);

                    // Copy each worksheet from the source workbook into the merged workbook
                    foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                    {
                        // Add a copy of the worksheet to the merged workbook by sheet name
                        mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                    }
                }

                // Set the author of the merged workbook to the current user name
                mergedWorkbook.Settings.Author = Environment.UserName;

                // Save the merged workbook with the updated author property
                mergedWorkbook.Save("MergedWorkbook_WithAuthor.xlsx");
                Console.WriteLine("Merge completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}