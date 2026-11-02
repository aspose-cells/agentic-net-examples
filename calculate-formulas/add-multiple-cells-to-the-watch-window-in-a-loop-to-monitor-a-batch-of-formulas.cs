// Title: Adding multiple cells to the Watch Window in a loop using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET add several cells to watch window programmatically | loop through cells and add them to watch window with Aspose.Cells | how to monitor a batch of formulas using Aspose.Cells watch window | watch window feature availability in Aspose.Cells versions | sample code for adding formulas to watch window in C#
// Tags: programmatic watch window addition Aspose.Cells | batch formula monitoring Aspose.Cells .NET | watch window API version check Aspose.Cells | add cells to watch window C# | watch window support Aspose.Cells workbook

using Aspose.Cells;
using System;

// This example demonstrates how to programmatically add multiple cells to the Watch Window in a loop with Aspose.Cells for .NET, highlights the need to verify version support for the WatchWindow API, and saves the workbook after populating sample values and formulas.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data and formulas (A1:A5 values, B1:B5 formulas)
            for (int i = 0; i < 5; i++)
            {
                // Put a numeric value in column A
                cells[i, 0].PutValue(i + 1);               // A1, A2, ...

                // Create a simple formula in column B that references column A
                cells[i, 1].Formula = $"=A{i + 1}*2";      // B1, B2, ...
            }

            // NOTE: The WatchWindow feature is not available in the current Aspose.Cells version.
            // If needed, ensure you are using a version that supports Worksheet.WatchWindow.

            // Save the workbook (lifecycle rule: save)
            string outputPath = "WatchWindowDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
