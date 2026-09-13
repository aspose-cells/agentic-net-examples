// Title: How to freeze the top three rows of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates an Excel workbook with Aspose.Cells, fills it with sample data, and freezes rows 1‑3 while keeping columns scrollable. | Demonstrate using Worksheet.FreezePanes to lock the first three rows in a .xlsx file with Aspose.Cells for .NET. | Show how to apply FreezePanes(3,0,0,0) and then save the workbook so header rows stay visible during navigation.
// Common Searches: Aspose.Cells C# freeze first three rows of worksheet example | How to keep header rows static while scrolling in Excel using Aspose.Cells .NET | Worksheet.FreezePanes parameters to freeze top rows in C# Aspose.Cells
// Tags: Aspose.Cells FreezePanes top rows | C# freeze header rows Excel | Aspose.Cells worksheet freeze rows example | Excel .xlsx freeze panes using Aspose.Cells | programmatic row freeze Aspose.Cells .NET

using System;
using Aspose.Cells;
using System.IO;

// The example creates a new workbook, populates cells with sample data, freezes the first three rows using Worksheet.FreezePanes(3,0,0,0) to keep header rows visible, and saves the file as FrozenTopThreeRows.xlsx.
class FreezeTopRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook with a default worksheet

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Populate some data to visualize the freeze effect
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top three rows (rows 0,1,2) while keeping all columns scrollable
            // FreezePanes(row, column, totalRows, totalColumns)
            // Setting totalRows and totalColumns to 0 freezes everything above 'row' and left of 'column'
            sheet.FreezePanes(3, 0, 0, 0); // freezes rows 0-2

            // Save the workbook to a file
            string outputPath = "FrozenTopThreeRows.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
