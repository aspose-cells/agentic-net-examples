// Title: Generate an Excel workbook with sample data and freeze the first column using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to freeze column A while creating a new workbook | Aspose.Cells create workbook with header row and sample rows in C# | Freeze first column in Excel using Aspose.Cells FreezePanes method example | Saving a workbook as .xlsx after applying FreezePanes with Aspose.Cells .NET | Add sample data to worksheet and lock first column using Aspose.Cells C# code
// Tags: freeze panes functionality Aspose.Cells C# | create workbook with sample rows Aspose.Cells | add header row to worksheet Aspose.Cells | save workbook as xlsx Aspose.Cells .NET | lock first column Excel Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program creates a new Workbook, writes a header row (ID, Name, Score) and three data rows, freezes column A via FreezePanes(0,1,0,0), saves the file as SampleData_FrozenColumn.xlsx, and outputs the full path while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Add sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Freeze the first column (column A)
            // Parameters: row index, column index, total rows to freeze, total columns to freeze
            // Setting row = 0 (no rows frozen), column = 1 (freeze column A), totals = 0
            sheet.FreezePanes(0, 1, 0, 0);

            // Define output file path
            string outputPath = "SampleData_FrozenColumn.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
