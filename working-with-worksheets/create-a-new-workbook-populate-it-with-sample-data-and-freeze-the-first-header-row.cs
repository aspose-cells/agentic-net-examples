// Title: Create a new Excel workbook, insert sample data rows, and freeze the header row using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a header row and several data rows, then freezes the first row. | Demonstrate the use of the FreezePanes method to lock the top header row after populating a worksheet with sample values.
// Common Searches: asp.net freeze first row in Excel with Aspose.Cells C# example | how to programmatically add header and data rows then freeze pane using Aspose.Cells | C# code to create an .xlsx file with sample data and freeze the header row
// Tags: Aspose.Cells create workbook with sample data | Aspose.Cells FreezePanes header row | C# populate worksheet cells programmatically | save Excel file as .xlsx using Aspose.Cells | initialize worksheet and lock top row

using Aspose.Cells;
using System;
using System.IO;

// C# program that creates a new workbook, writes a header and three data rows, freezes the first row using FreezePanes, and saves the file as SampleData.xlsx with Aspose.Cells for .NET.
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

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Freeze the first header row (freeze rows above row 2)
            // Parameters: row, column, totalRowsToFreeze, totalColumnsToFreeze
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output path
            string outputPath = "SampleData.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
