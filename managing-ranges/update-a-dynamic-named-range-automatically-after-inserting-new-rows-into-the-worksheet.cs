// Title: Automatically update a dynamic named range after inserting rows with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a dynamic named range in column A using an INDEX‑COUNTA formula, insert rows, add data, and call RefreshDynamicArrayFormulas to keep the range address current before saving the workbook.
// Keywords: Aspose.Cells C# dynamic named range | RefreshDynamicArrayFormulas .NET | update named range after row insertion | INDEX COUNTA named range formula | insert rows Aspose.Cells | auto‑expand named range | Excel automation Aspose.Cells
// Common Searches: Aspose.Cells keep dynamic named range up to date after inserting rows | RefreshDynamicArrayFormulas usage in C# | auto expand named range with INDEX and COUNTA in Aspose.Cells | C# insert rows and update named range in Excel workbook | Aspose.Cells dynamic range after row insertion
// Developer Intent: Make a dynamic named range automatically expand to include newly inserted rows in a worksheet.
// Use Cases: Define a dynamic named range that grows with non‑empty cells in a column using INDEX and COUNTA. | Insert multiple rows at a specific position and have the named range adjust without manual recalculation. | Refresh dynamic array formulas to obtain the updated range address after structural changes.
// AI Prompts: Write C# code that creates a dynamic named range with INDEX/COUNTA in Aspose.Cells and updates it after inserting rows. | Explain when and why RefreshDynamicArrayFormulas should be called in an Aspose.Cells workbook. | Show how to retrieve the updated address of a dynamic named range after adding rows to a worksheet.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a dynamic named range in column A using an INDEX‑COUNTA formula, insert rows, add data, and call RefreshDynamicArrayFormulas to keep the range address current before saving the workbook.
class UpdateDynamicNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate column A with initial data (5 rows)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Define a dynamic named range that expands with the number of non‑empty rows in column A
            // Formula: =Sheet1!$A$1:INDEX(Sheet1!$A:$A, COUNTA(Sheet1!$A:$A))
            int nameIdx = wb.Worksheets.Names.Add("MyDynamicRange");
            Name dynName = wb.Worksheets.Names[nameIdx];
            dynName.RefersTo = "=Sheet1!$A$1:INDEX(Sheet1!$A:$A, COUNTA(Sheet1!$A:$A))";

            // Display the initial range address
            AsposeRange initRange = dynName.GetRange();
            Console.WriteLine($"Initial range address: {initRange.Address}");

            // Insert three rows at index 2 (third row) and update references
            ws.Cells.InsertRows(2, 3, true);

            // Add data into the newly inserted rows
            cells[2, 0].PutValue("Inserted 1");
            cells[3, 0].PutValue("Inserted 2");
            cells[4, 0].PutValue("Inserted 3");

            // Refresh dynamic array formulas (required after row insertion)
            wb.RefreshDynamicArrayFormulas(true);

            // Retrieve and display the updated range address
            AsposeRange updatedRange = dynName.GetRange();
            Console.WriteLine($"Updated range address after insertion: {updatedRange.Address}");

            // Save the workbook
            wb.Save("DynamicNamedRangeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
