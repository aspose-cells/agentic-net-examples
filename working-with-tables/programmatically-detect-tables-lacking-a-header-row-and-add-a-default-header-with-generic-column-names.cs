// Title: Programmatically add default column headers to Excel tables lacking a header row using Aspose.Cells for .NET
// Common Searches: c# aspnet add generic column names to Excel ListObject tables that have hidden headers | detect Excel tables without visible header row and insert Column1 Column2 using Aspose.Cells | show header row for tables in a workbook and assign default names with Aspose.Cells C# | iterate worksheets and listobjects to set missing table headers in .NET | how to automatically create column headers for tables that hide the header in Aspose.Cells
// Tags: add default column headers to ListObject tables | detect tables with hidden header rows using Aspose.Cells | show header row and set generic column names in C# | iterate worksheets to fix missing table headers | Aspose.Cells table header row handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook, loops through each worksheet and its ListObject tables, reveals any hidden header rows, assigns generic names like Column1, Column2, etc., and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process each table (ListObject) in the worksheet
                foreach (ListObject table in sheet.ListObjects)
                {
                    // If header row is hidden, show it and assign generic column names
                    if (!table.ShowHeaderRow)
                    {
                        table.ShowHeaderRow = true;

                        int columnCount = table.ListColumns.Count;
                        for (int i = 0; i < columnCount; i++)
                        {
                            string headerName = $"Column{i + 1}";
                            table.ListColumns[i].Name = headerName;
                        }
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
