// Title: Convert an Excel table to a regular range while keeping only the header row formatting and save the workbook as ODS using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert ListObject to range preserving header style | How to retain table header formatting after converting to range with Aspose.Cells | Save Excel workbook as ODS after table-to-range conversion in .NET | C# code to export a table as ODS while keeping header formatting using Aspose.Cells | Aspose.Cells example for converting Excel table to range and exporting to ODS
// Tags: ListObject to range conversion Aspose.Cells | header style preservation Aspose.Cells | export workbook to ODS Aspose.Cells | Excel table conversion C# Aspose.Cells | retain table header formatting .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an Excel file, captures the styles of the table's header cells, converts the ListObject to a normal range, reapplies the saved header styles, and saves the workbook as an ODS file using Aspose.Cells for .NET.
class TableToRangeConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.ods";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook (Excel file)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }

            // Get the first table on the sheet
            ListObject table = sheet.ListObjects[0];

            // Calculate column count from start and end columns
            int headerColumnCount = table.EndColumn - table.StartColumn + 1;

            // Preserve header row formatting by storing each header cell's style
            Style[] headerStyles = new Style[headerColumnCount];
            for (int col = 0; col < headerColumnCount; col++)
            {
                // Header cells are located at the first row of the table range
                Cell headerCell = sheet.Cells[table.StartRow, table.StartColumn + col];
                headerStyles[col] = headerCell.GetStyle(); // GetStyle returns a copy
            }

            // Convert the table to a normal range
            table.ConvertToRange();

            // Reapply the saved header styles to the first row (header) of the new range
            for (int col = 0; col < headerColumnCount; col++)
            {
                Cell headerCell = sheet.Cells[table.StartRow, table.StartColumn + col];
                headerCell.SetStyle(headerStyles[col]);
            }

            // Save the workbook as ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
