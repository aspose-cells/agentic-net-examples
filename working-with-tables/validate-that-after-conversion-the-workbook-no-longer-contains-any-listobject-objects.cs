// Title: C# – Verify No ListObjects Remain After Converting Tables to Ranges with Aspose.Cells
// Description: Loads an Excel workbook, iterates through each worksheet, converts every ListObject (table) to a regular range, confirms the ListObjects collection is empty, and saves the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# convert ListObject to range | remove Excel tables Aspose.Cells | validate ListObjects count | Aspose.Cells ListObject verification | convert tables to ranges .NET
// Common Searches: Aspose.Cells convert all tables to ranges | how to delete ListObjects after ConvertToRange | check for remaining ListObjects in workbook C# | ensure no Excel tables exist after conversion Aspose.Cells | bulk remove ListObjects from worksheets
// Developer Intent: Confirm that after converting every ListObject to a range, the workbook contains zero ListObject objects.
// Use Cases: Strip tables before exporting to CSV to avoid table‑specific formatting. | Validate workbook cleanliness after bulk table removal for downstream processing. | Increase compatibility with older Excel versions that lack ListObject support.
// AI Prompts: Generate C# code using Aspose.Cells that converts all ListObjects in a workbook to ranges and asserts the ListObjects collection is empty. | Write a unit test that loads a sample workbook, runs the conversion loop, and verifies Worksheet.ListObjects.Count equals zero for every sheet. | Explain how ConvertToRange updates the ListObjects collection and recommend the safest iteration pattern for removing all tables.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel workbook, iterates through each worksheet, converts every ListObject (table) to a regular range, confirms the ListObjects collection is empty, and saves the updated file using Aspose.Cells for .NET.
class ValidateNoListObjects
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Process each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Convert all ListObjects (tables) in the worksheet to regular ranges
            // Loop until the collection is empty because ConvertToRange removes the object
            while (sheet.ListObjects.Count > 0)
            {
                ListObject table = sheet.ListObjects[0];
                table.ConvertToRange();
            }

            // Verify that no ListObjects remain
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' contains no ListObjects after conversion.");
            }
            else
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' still contains {sheet.ListObjects.Count} ListObjects.");
            }
        }

        // Save the modified workbook (replace with your desired output file)
        workbook.Save("output.xlsx");
    }
}
