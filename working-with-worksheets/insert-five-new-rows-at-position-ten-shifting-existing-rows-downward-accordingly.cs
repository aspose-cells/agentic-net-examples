// Title: Insert five rows at row 10 in an Aspose.Cells worksheet with C# while shifting existing rows down
// AI Prompts: Write C# code that uses Aspose.Cells to insert a block of five rows at index 9 in a worksheet and then saves the workbook. | Show how to call Worksheet.Cells.InsertRows to add multiple rows at a specific position and keep the original data intact.
// Common Searches: Aspose.Cells C# insert multiple rows at a specific row index | How to add five rows at row 10 in an Excel file using Aspose.Cells .NET | InsertRows method example for shifting rows down in Aspose.Cells | C# code to insert rows and preserve existing worksheet data with Aspose.Cells | Insert rows at position 10 in workbook using Aspose.Cells API
// Tags: Aspose.Cells InsertRows API | C# add multiple rows Excel worksheet | shift existing rows down Aspose.Cells | insert rows at specific index .NET | Excel file row insertion using Aspose

using System;
using Aspose.Cells;

// // Demonstrates creating a workbook, populating the first column with sample data, inserting five rows at the 10th position (zero‑based index 9) which shifts existing rows downward, and saving the result to InsertRowsResult.xlsx.
class InsertRowsExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data to illustrate the shift after insertion
        for (int i = 0; i < 15; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Insert five rows at position ten (zero‑based index 9)
        worksheet.Cells.InsertRows(9, 5);

        // Save the workbook
        workbook.Save("InsertRowsResult.xlsx");
    }
}
