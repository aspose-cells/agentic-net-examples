// Title: Validate Detail Worksheet Row Count Against Source Collection with Aspose.Cells for .NET
// Description: C# example that creates a workbook, writes a List<string[]> to the first worksheet named "Detail", then checks Cells.Rows.Count against the source list size. The script prints a pass/fail message and saves the file, demonstrating how to ensure all records are written.
// Keywords: Aspose.Cells row count validation | C# Excel row verification | worksheet rows vs list size | Aspose.Cells .NET example | Excel data export validation | smart markers row check | global developers
// Common Searches: Aspose.Cells verify worksheet row count | C# compare Cells.Rows.Count with collection size | how to confirm all rows are written to Excel using Aspose | validate Excel detail sheet record count .NET | row count mismatch handling Aspose.Cells
// Developer Intent: Confirm that the detail sheet contains exactly the same number of rows as the source collection before saving the workbook.
// Use Cases: Automated reporting pipelines that need to guarantee every record appears in the Excel detail section. | Data migration scripts where missing rows must be detected early. | Quality‑assurance checks in batch Excel generation to prevent incomplete files.
// AI Prompts: Write C# code that populates an Aspose.Cells worksheet from a List<string[]> and asserts Cells.Rows.Count equals the list count, logging success or failure. | Show error‑handling patterns for a row‑count mismatch when exporting data with Aspose.Cells, including throwing an exception or returning a status code. | Create a reusable C# method that accepts a Worksheet and IEnumerable<T>, writes the data, validates the row count, and returns a boolean result.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// C# example that creates a workbook, writes a List<string[]> to the first worksheet named "Detail", then checks Cells.Rows.Count against the source list size. The script prints a pass/fail message and saves the file, demonstrating how to ensure all records are written.
class Program
{
    static void Main()
    {
        // Sample source collection representing rows to be written to the worksheet
        List<string[]> sourceData = new List<string[]>
        {
            new [] { "ID", "Name", "Qty" },
            new [] { "1", "Apple", "10" },
            new [] { "2", "Banana", "20" },
            new [] { "3", "Orange", "15" }
        };

        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Use the first worksheet as the detail sheet
        Worksheet detailSheet = workbook.Worksheets[0];
        detailSheet.Name = "Detail";

        // Populate the worksheet with data from the source collection
        for (int i = 0; i < sourceData.Count; i++)
        {
            string[] row = sourceData[i];
            for (int j = 0; j < row.Length; j++)
            {
                detailSheet.Cells[i, j].PutValue(row[j]);
            }
        }

        // Validate that the number of rows in the worksheet matches the source collection count
        int expectedRowCount = sourceData.Count;
        int actualRowCount = detailSheet.Cells.Rows.Count; // RowCollection.Count

        if (actualRowCount == expectedRowCount)
        {
            Console.WriteLine($"Validation passed: {actualRowCount} rows present as expected.");
        }
        else
        {
            Console.WriteLine($"Validation failed: expected {expectedRowCount} rows but found {actualRowCount}.");
        }

        // Save the workbook (save rule)
        workbook.Save("DetailValidation.xlsx");
    }
}
