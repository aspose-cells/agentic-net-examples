// Title: Create a new Excel workbook, add sample data, and freeze the header row with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# freeze top row while scrolling | Create Excel file with sample data using Aspose.Cells .NET | How to use FreezePanes method in Aspose.Cells to lock header row | Saving workbook as .xlsx after adding data with Aspose.Cells | Initialize new workbook and write data to cells in C# Aspose.Cells
// Tags: freeze panes API Aspose.Cells | sample data insertion Aspose.Cells | xlsx workbook saving Aspose.Cells | new workbook creation Aspose.Cells | header row freezing Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses the first Worksheet, writes a header row and three rows of sample data, freezes the top header row using the FreezePanes method, ensures the output directory exists, saves the file as SampleData.xlsx, and prints a success message while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate sample data
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Freeze the first header row (freeze rows above row index 1)
            // Use the overload with four parameters (row, column, totalRows, totalColumns)
            sheet.FreezePanes(1, 0, 0, 0);

            // Define output file path
            string outputPath = "SampleData.xlsx";

            // Ensure the output directory exists (if a directory part is present)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
