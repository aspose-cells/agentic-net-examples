// Title: Clone an Aspose.Cells workbook and retain all named ranges in C#
// Description: Demonstrates how to create a source workbook, add named ranges via Worksheets.Names and Range.Name, clone the workbook with CopyOptions.CopyNames enabled, verify that each name and its RefersTo reference are preserved, and save both workbooks.
// Keywords: Aspose.Cells clone workbook C# | preserve named ranges Aspose.Cells | CopyOptions CopyNames true | verify workbook copy named ranges | Aspose.Cells range name duplication check
// Common Searches: clone Aspose.Cells workbook with named ranges | copy workbook keeping range names .NET | how to verify named ranges after workbook copy | Aspose.Cells CopyOptions CopyNames example | C# preserve named ranges when duplicating Excel file
// Developer Intent: The developer needs to duplicate an Excel workbook while ensuring that every defined name (named range) is copied exactly and can be programmatically validated.
// Use Cases: Generate user‑specific reports from a template that contains named ranges without breaking formulas. | Automate testing of workbook cloning logic by confirming that all Name objects survive the copy operation. | Create backup copies of workbooks where named ranges must remain intact for downstream processing.
// AI Prompts: Provide C# code that clones an Aspose.Cells workbook with CopyOptions.CopyNames set to true and checks that all named ranges match the source. | Explain step‑by‑step how to compare the Names collections of two workbooks to confirm identical RefersTo references. | Describe how named ranges created via Worksheets.Names differ from those set with Range.Name when using Workbook.Copy.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCloneNamedRanges
{
    // Demonstrates how to create a source workbook, add named ranges via Worksheets.Names and Range.Name, clone the workbook with CopyOptions.CopyNames enabled, verify that each name and its RefersTo reference are preserved, and save both workbooks.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and add named ranges ----------
                Workbook sourceWorkbook = new Workbook();
                Worksheet sheet = sourceWorkbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate some cells
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["B3"].PutValue(40);

                // Create named range using the Names collection
                int idx1 = sourceWorkbook.Worksheets.Names.Add("Numbers");
                sourceWorkbook.Worksheets.Names[idx1].RefersTo = "=Data!$A$2:$A$3";

                // Create another named range using the Range.Name property
                AsposeRange rng = sheet.Cells.CreateRange("B2:B3");
                rng.Name = "Values";

                // ---------- Clone the workbook preserving named ranges ----------
                Workbook clonedWorkbook = new Workbook();
                CopyOptions options = new CopyOptions { CopyNames = true };
                clonedWorkbook.Copy(sourceWorkbook, options);

                // ---------- Verify that all named ranges are present in the clone ----------
                Console.WriteLine("Verification of named ranges in the cloned workbook:");
                foreach (Name srcName in sourceWorkbook.Worksheets.Names)
                {
                    // Retrieve the same name from the cloned workbook
                    Name destName = clonedWorkbook.Worksheets.Names[srcName.Text];

                    if (destName != null && destName.RefersTo == srcName.RefersTo)
                    {
                        Console.WriteLine($"- Name '{srcName.Text}' copied successfully. RefersTo: {destName.RefersTo}");
                    }
                    else
                    {
                        Console.WriteLine($"- Name '{srcName.Text}' was NOT copied correctly.");
                    }
                }

                // Optional: Save both workbooks to verify manually
                string sourcePath = "SourceWorkbook.xlsx";
                string clonedPath = "ClonedWorkbook.xlsx";

                // Ensure we can write to the directory
                string directory = Path.GetDirectoryName(Path.GetFullPath(sourcePath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx);
                clonedWorkbook.Save(clonedPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
