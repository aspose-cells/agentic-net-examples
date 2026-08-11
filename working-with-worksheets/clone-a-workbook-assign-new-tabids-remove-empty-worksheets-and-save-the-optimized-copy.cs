// Title: Clone a Workbook, Remove Empty Sheets, Assign New TabIds, and Save Optimized Copy with Aspose.Cells for .NET (C#)
// Description: Loads a source XLSX, creates a fresh workbook, copies only worksheets that contain data, assigns sequential TabId values starting at 100, guarantees at least one sheet, and saves the streamlined file as optimized_output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy worksheet C# | Aspose.Cells remove empty worksheets | Aspose.Cells set TabId | clone workbook Aspose | optimize Excel file .NET | filter blank sheets Aspose.Cells | Workbook optimization C# | Aspose.Cells SaveFormat Xlsx | Excel workbook cleanup | C# Excel performance
// Common Searches: copy only non‑empty worksheets Aspose.Cells C# | assign custom TabId values after copying worksheets Aspose | remove default and empty sheets with Aspose.Cells .NET | optimize large Excel workbook by stripping blank sheets | Aspose.Cells clone workbook and reindex tabs
// Developer Intent: Generate a new workbook that contains only populated worksheets from an existing file, give each sheet a fresh TabId, and persist the result as a lightweight XLSX.
// Use Cases: Create a slim version of a massive workbook for distribution by discarding empty tabs. | Prepare files for downstream systems that require worksheets to have TabIds in a specific numeric range. | Automate template sanitization before publishing, keeping only sheets with actual data.
// AI Prompts: Write C# code with Aspose.Cells that copies only data‑bearing worksheets from a source workbook, assigns TabId values starting at 200, and saves the result. | Provide a reusable method that accepts a source file path, removes empty worksheets, reassigns sequential TabIds, and returns the optimized Workbook object. | Explain how to ensure a newly created workbook always contains at least one worksheet after empty‑sheet removal using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookOptimizationDemo
{
    // Loads a source XLSX, creates a fresh workbook, copies only worksheets that contain data, assigns sequential TabId values starting at 100, guarantees at least one sheet, and saves the streamlined file as optimized_output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "source.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWb = new Workbook(sourcePath);

                // Create a new empty workbook
                Workbook optimizedWb = new Workbook();

                // Remove the default worksheet if present
                if (optimizedWb.Worksheets.Count > 0)
                {
                    optimizedWb.Worksheets.RemoveAt(0);
                }

                // Iterate through each worksheet in the source workbook
                for (int i = 0; i < sourceWb.Worksheets.Count; i++)
                {
                    Worksheet srcSheet = sourceWb.Worksheets[i];

                    // Determine if the worksheet contains any data
                    bool hasData = srcSheet.Cells.MaxDataRow >= 0 && srcSheet.Cells.MaxDataColumn >= 0;

                    if (hasData)
                    {
                        // Copy the non‑empty worksheet to the optimized workbook
                        int newIndex = optimizedWb.Worksheets.AddCopy(srcSheet.Name);
                        Worksheet destSheet = optimizedWb.Worksheets[newIndex];

                        // Assign a new TabId (e.g., sequential starting from 100)
                        destSheet.TabId = 100 + newIndex;
                    }
                }

                // Ensure at least one worksheet exists
                if (optimizedWb.Worksheets.Count == 0)
                {
                    optimizedWb.Worksheets.Add();
                    optimizedWb.Worksheets[0].TabId = 100;
                }

                // Save the optimized workbook
                optimizedWb.Save("optimized_output.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Optimized workbook saved as optimized_output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
