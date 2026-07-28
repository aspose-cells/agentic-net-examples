// Title: Clone a Workbook with Named Ranges Using Aspose.Cells for .NET and Verify Preservation
// Description: Demonstrates how to create a source workbook with multiple named ranges, clone it with CopyOptions.CopyNames enabled, iterate through the original names to confirm each exists in the clone with the same RefersTo reference, report mismatches, and optionally save the cloned file.
// Keywords: Aspose.Cells clone workbook C# | CopyOptions CopyNames true | preserve named ranges Aspose.Cells | verify named range after workbook copy | duplicate workbook with named ranges .NET | Aspose.Cells named range verification
// Common Searches: How to copy a workbook with named ranges using Aspose.Cells for .NET | Aspose.Cells preserve named ranges when cloning a workbook | CopyOptions.CopyNames example C# | Check named range references after workbook duplication Aspose.Cells | Validate named ranges in cloned Excel file using Aspose.Cells
// Developer Intent: Clone an existing workbook while automatically copying all defined named ranges and programmatically confirm that each range in the clone matches the original.
// Use Cases: Generate personalized reports from a master template without losing named range definitions. | Automate creation of multiple workbooks for batch processing while keeping data‑extraction formulas intact. | Run integration tests that duplicate a workbook and verify that named ranges remain consistent between source and clone.
// AI Prompts: Write C# code with Aspose.Cells that clones a workbook, copies all named ranges, and logs any missing or mismatched ranges. | Show an example using CopyOptions.CopyNames = true to duplicate a workbook and then compare each Name.RefersTo between source and clone. | Create a unit test in C# that asserts the cloned workbook contains the same named ranges as the original workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a source workbook with multiple named ranges, clone it with CopyOptions.CopyNames enabled, iterate through the original names to confirm each exists in the clone with the same RefersTo reference, report mismatches, and optionally save the cloned file.
    public class CloneWorkbookWithNamedRangesDemo
    {
        public static void Run()
        {
            try
            {
                // ---------- Create source workbook and define named ranges ----------
                Workbook sourceWorkbook = new Workbook();

                // First worksheet with a named range
                Worksheet sheet1 = sourceWorkbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("First Value");
                int nameIndex1 = sourceWorkbook.Worksheets.Names.Add("FirstRange");
                sourceWorkbook.Worksheets.Names[nameIndex1].RefersTo = "=Sheet1!$A$1";

                // Second worksheet with another named range
                Worksheet sheet2 = sourceWorkbook.Worksheets.Add("Data");
                sheet2.Cells["C3"].PutValue("Second Value");
                int nameIndex2 = sourceWorkbook.Worksheets.Names.Add("SecondRange");
                sourceWorkbook.Worksheets.Names[nameIndex2].RefersTo = "=Data!$C$3";

                // ---------- Clone the workbook while preserving named ranges ----------
                Workbook clonedWorkbook = new Workbook();
                CopyOptions copyOptions = new CopyOptions
                {
                    CopyNames = true // Ensure named ranges are copied
                };
                clonedWorkbook.Copy(sourceWorkbook, copyOptions);

                // ---------- Verify that all named ranges are present in the clone ----------
                bool allNamesPreserved = true;
                foreach (Name sourceName in sourceWorkbook.Worksheets.Names)
                {
                    // Retrieve the corresponding name in the cloned workbook by its text
                    Name clonedName = clonedWorkbook.Worksheets.Names[sourceName.Text];

                    // Check existence and reference equality
                    if (clonedName == null || clonedName.RefersTo != sourceName.RefersTo)
                    {
                        allNamesPreserved = false;
                        Console.WriteLine($"Missing or mismatched named range: {sourceName.Text}");
                    }
                    else
                    {
                        Console.WriteLine($"Named range '{sourceName.Text}' copied successfully: {clonedName.RefersTo}");
                    }
                }

                Console.WriteLine(allNamesPreserved
                    ? "All named ranges were preserved in the cloned workbook."
                    : "Some named ranges were not preserved.");

                // ---------- Optional: Save the cloned workbook for manual inspection ----------
                clonedWorkbook.Save("ClonedWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CloneWorkbookWithNamedRangesDemo.Run();
        }
    }
}
