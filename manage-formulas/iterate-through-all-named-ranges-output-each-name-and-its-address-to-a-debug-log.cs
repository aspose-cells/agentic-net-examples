// Title: C# – List all named ranges in an Aspose.Cells workbook and log their addresses
// Description: Creates a workbook, defines sample named ranges, retrieves every named range with Workbook.Worksheets.GetNamedRanges(), and writes each range's Name and RefersTo address to the debug output using Debug.WriteLine. Includes handling for workbooks without named ranges.
// Keywords: Aspose.Cells C# list named ranges | GetNamedRanges | Debug.WriteLine named range | range.RefersTo address | enumerate named ranges Aspose.Cells | C# workbook named range logging
// Common Searches: how to enumerate named ranges Aspose.Cells C# | Aspose.Cells get all named ranges and addresses | debug log named range name and address Aspose.Cells | C# list workbook named ranges Aspose.Cells | retrieve named range RefersTo property
// Developer Intent: Extract every named range from a workbook and output its name and cell reference to the debug console.
// Use Cases: Verify that required named ranges exist before running calculations | Generate a quick reference sheet of all defined ranges for documentation | Troubleshoot formula or data‑validation issues by logging range details
// AI Prompts: Generate C# code with Aspose.Cells that exports all named range names and addresses to a CSV file. | Show how to filter named ranges by a prefix and log only the matching ones using Aspose.Cells. | Provide a robust example that safely handles a workbook with no named ranges without throwing exceptions.

using System;
using System.Diagnostics;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeDemo
{
    // Creates a workbook, defines sample named ranges, retrieves every named range with Workbook.Worksheets.GetNamedRanges(), and writes each range's Name and RefersTo address to the debug output using Debug.WriteLine. Includes handling for workbooks without named ranges.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // ------------------------------------------------------------
                // Sample data: create a couple of named ranges for demonstration
                // ------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Item1");
                sheet.Cells["A2"].PutValue("Item2");
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);

                // Define named ranges
                sheet.Cells.CreateRange("A1:A2").Name = "ItemNames";
                sheet.Cells.CreateRange("B1:B2").Name = "ItemValues";

                // ------------------------------------------------------------
                // Iterate through all named ranges and write name & address to debug log
                // ------------------------------------------------------------
                // Get all pre‑defined named ranges in the workbook
                AsposeRange[] namedRanges = workbook.Worksheets.GetNamedRanges();

                if (namedRanges != null && namedRanges.Length > 0)
                {
                    foreach (AsposeRange range in namedRanges)
                    {
                        // Output the name and its address using Debug.WriteLine
                        Debug.WriteLine($"Name: {range.Name}, Address: {range.RefersTo}");
                    }
                }
                else
                {
                    Debug.WriteLine("No named ranges found in the workbook.");
                }

                // ------------------------------------------------------------
                // (Optional) Save the workbook if needed
                // ------------------------------------------------------------
                // workbook.Save("NamedRangesDemo.xlsx");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Debug.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
