// Title: Copy specific worksheets from two Excel workbooks into a new workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to copy the "Report" sheet from SourceWorkbook1.xlsx and the third worksheet from SourceWorkbook2.xlsx into a freshly created workbook, keeping the original sheet names. | Create a .NET program that merges selected sheets from multiple workbooks into one file, includes checks for missing source files and save errors, and utilizes the Worksheet.AddCopy method for copying.
// Common Searches: asp.net copy worksheet by name from one Excel file to another using Aspose.Cells | how to add a copy of a worksheet from a second workbook in C# with Aspose.Cells | copy third sheet from an Excel workbook to a new workbook Aspose.Cells example | merge selected sheets from several Excel files into one workbook C# Aspose.Cells | error handling when copying worksheets between workbooks with Aspose.Cells
// Tags: Worksheet.AddCopy copy worksheet Aspose.Cells | copy specific worksheet by name C# | merge selected Excel sheets Aspose.Cells | copy third worksheet index Aspose.Cells | combine multiple workbooks preserving sheet names

using Aspose.Cells;
using System;
using System.IO;

// The example loads two source Excel files, copies the "Report" sheet from the first workbook and the third sheet (index 2) from the second workbook into a newly created workbook using Worksheet.AddCopy, retains the original sheet names, and saves the result as CombinedWorkbook.xlsx. It includes file‑existence checks and comprehensive error handling for copy and save operations.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to source workbooks
            string sourcePath1 = "SourceWorkbook1.xlsx";
            string sourcePath2 = "SourceWorkbook2.xlsx";

            // Verify that source files exist
            if (!File.Exists(sourcePath1) || !File.Exists(sourcePath2))
            {
                Console.WriteLine("One or both source files are missing.");
                return;
            }

            // Load source workbooks
            Workbook sourceWorkbook1 = new Workbook(sourcePath1);
            Workbook sourceWorkbook2 = new Workbook(sourcePath2);

            // Create target workbook and remove the default sheet
            Workbook targetWorkbook = new Workbook();
            targetWorkbook.Worksheets.Clear();

            // -------------------------------------------------
            // Copy specific worksheet from the first source
            // -------------------------------------------------
            Worksheet sourceSheet1 = sourceWorkbook1.Worksheets["Report"];
            if (sourceSheet1 != null)
            {
                try
                {
                    // Add a copy of the source sheet to the target workbook using its name
                    int newIndex = targetWorkbook.Worksheets.AddCopy(sourceSheet1.Name);
                    // Optionally rename the copied sheet (keeps original name here)
                    targetWorkbook.Worksheets[newIndex].Name = sourceSheet1.Name;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy sheet 'Report' from first workbook: {ex.Message}");
                }
            }

            // -------------------------------------------------
            // Copy specific worksheet from the second source
            // -------------------------------------------------
            if (sourceWorkbook2.Worksheets.Count > 2)
            {
                try
                {
                    Worksheet sourceSheet2 = sourceWorkbook2.Worksheets[2]; // third sheet (zero‑based index)

                    // Add a copy of the source sheet to the target workbook using its name
                    int newIndex = targetWorkbook.Worksheets.AddCopy(sourceSheet2.Name);
                    // Optionally rename the copied sheet
                    targetWorkbook.Worksheets[newIndex].Name = sourceSheet2.Name;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy third sheet from second workbook: {ex.Message}");
                }
            }

            // -------------------------------------------------
            // Save the combined workbook
            // -------------------------------------------------
            string outputPath = "CombinedWorkbook.xlsx";
            try
            {
                targetWorkbook.Save(outputPath);
                Console.WriteLine($"Combined workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save combined workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
