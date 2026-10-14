// Title: Create a new Excel workbook, add sample data, and freeze the first column using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a workbook, write a header row and several data rows, then freeze column A while keeping the rest of the sheet scrollable. | Show the exact FreezePanes parameters required to lock the first column in an Aspose.Cells worksheet and save the file as an XLSX document. | Extend the example to also freeze the top header row together with the first column using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# freeze first column without freezing rows | C# example to create Excel file with sample data and freeze column A using Aspose.Cells | How to use FreezePanes in Aspose.Cells to lock column A and keep other columns scrollable | Saving a new workbook with header and data rows to XLSX using Aspose.Cells .NET | Freeze panes for first column only in Aspose.Cells C# code
// Tags: freeze column A Aspose.Cells | populate worksheet with sample data C# | create and save new workbook Aspose.Cells | apply FreezePanes to lock first column .NET | add header row and data rows Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook, writes a header row and three data rows, freezes column A using FreezePanes, and saves the file as SampleData.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data (header + a few rows)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

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
            // FreezePanes(totalRows, totalColumns, row, column)
            // To freeze column A, set column index to 1 (column B becomes the first scrollable column)
            sheet.FreezePanes(0, 1, 0, 1);

            // Define output file path
            string outputPath = "SampleData.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
