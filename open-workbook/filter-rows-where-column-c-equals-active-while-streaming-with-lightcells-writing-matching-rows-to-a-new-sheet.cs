using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and populate source data
            Workbook workbook = new Workbook();
            Worksheet source = workbook.Worksheets[0];
            source.Name = "Source";

            // Header row
            source.Cells["A1"].PutValue("ID");
            source.Cells["B1"].PutValue("Name");
            source.Cells["C1"].PutValue("Status");

            // Sample rows
            source.Cells["A2"].PutValue(1);
            source.Cells["B2"].PutValue("Alice");
            source.Cells["C2"].PutValue("Active");

            source.Cells["A3"].PutValue(2);
            source.Cells["B3"].PutValue("Bob");
            source.Cells["C3"].PutValue("Inactive");

            source.Cells["A4"].PutValue(3);
            source.Cells["B4"].PutValue("Charlie");
            source.Cells["C4"].PutValue("Active");

            // Add a new worksheet that will receive the filtered rows
            int filteredSheetIndex = workbook.Worksheets.Add();
            Worksheet filtered = workbook.Worksheets[filteredSheetIndex];
            filtered.Name = "Filtered";

            // Copy header
            for (int col = 0; col <= source.Cells.MaxColumn; col++)
            {
                filtered.Cells[0, col].PutValue(source.Cells[0, col].Value);
            }

            // Copy rows where column C (index 2) equals "Active"
            int targetRow = 1; // start after header
            for (int row = 0; row <= source.Cells.MaxDataRow; row++)
            {
                var statusCell = source.Cells[row, 2];
                if (statusCell != null && statusCell.StringValue == "Active")
                {
                    for (int col = 0; col <= source.Cells.MaxColumn; col++)
                    {
                        filtered.Cells[targetRow, col].PutValue(source.Cells[row, col].Value);
                    }
                    targetRow++;
                }
            }

            // Save the workbook
            string outputPath = "FilteredResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}