// Title: Set sequential TabId for each worksheet, verify uniqueness, and save the workbook with Aspose.Cells for .NET
// Common Searches: C# assign incremental TabId to all worksheets using Aspose.Cells | how to ensure worksheet TabId values are unique before saving with Aspose.Cells | Aspose.Cells check for duplicate TabId across sheets in a workbook | update TabId property for multiple worksheets programmatically in .NET | save Excel file after validating worksheet TabId uniqueness with Aspose.Cells
// Tags: sequential worksheet TabId assignment Aspose.Cells | worksheet TabId uniqueness validation Aspose.Cells | duplicate TabId detection in workbook C# | Aspose.Cells save workbook after TabId verification | C# update worksheet TabId property

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file or creates a new workbook, iterates through each worksheet assigning sequential TabId values starting at 1, uses a HashSet to confirm that all TabId values are unique, aborts the save if duplicates are found, and otherwise saves the workbook to the specified output path while providing console logging and exception handling.
class Program
{
    static void Main(string[] args)
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample sheets
            if (File.Exists(inputPath))
            {
                Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");
                workbook = new Workbook();

                // Add a few sample worksheets
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");
                workbook.Worksheets.Add("Sheet4");
            }

            // Update TabId values to sequential numbers starting from 1
            Console.WriteLine("Updating TabId values to sequential numbers...");
            int nextId = 1;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.TabId = nextId;
                Console.WriteLine($"Worksheet \"{ws.Name}\" assigned TabId = {nextId}");
                nextId++;
            }

            // Verify uniqueness of TabId across all worksheets
            Console.WriteLine("Verifying uniqueness of TabId values...");
            HashSet<int> tabIdSet = new HashSet<int>();
            bool duplicateFound = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (!tabIdSet.Add(ws.TabId))
                {
                    Console.WriteLine($"Duplicate TabId detected: {ws.TabId} on worksheet \"{ws.Name}\"");
                    duplicateFound = true;
                }
            }

            if (duplicateFound)
            {
                Console.WriteLine("Error: Duplicate TabId values exist after assignment. Aborting save.");
                return;
            }
            else
            {
                Console.WriteLine("All TabId values are unique.");
            }

            // Save the workbook
            Console.WriteLine($"Saving workbook to \"{outputPath}\"...");
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
