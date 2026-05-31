using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

public class TableToRangeCopyDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet (source)
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate sample data for the table
            srcSheet.Cells["A1"].PutValue("ID");
            srcSheet.Cells["B1"].PutValue("Name");
            srcSheet.Cells["A2"].PutValue(1);
            srcSheet.Cells["B2"].PutValue("John");
            srcSheet.Cells["A3"].PutValue(2);
            srcSheet.Cells["B3"].PutValue("Mary");

            // Add a ListObject (table) covering the data range A1:B3
            int tableIdx = srcSheet.ListObjects.Add("A1", "B3", true);
            ListObject table = srcSheet.ListObjects[tableIdx];

            // Determine the size of the table
            int startRow = table.StartRow;
            int startCol = table.StartColumn;
            int rowCount = table.EndRow - table.StartRow + 1;
            int colCount = table.EndColumn - table.StartColumn + 1;

            // Create a Range object that represents the table area (before conversion)
            AsposeRange sourceRange = srcSheet.Cells.CreateRange(startRow, startCol, rowCount, colCount);

            // Convert the table to a normal range (the ListObject is removed)
            table.ConvertToRange();

            // Add a destination worksheet
            Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            destSheet.Name = "Destination";

            // Create a destination range with the same dimensions starting at A1
            AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, rowCount, colCount);

            // Copy the source range (now a plain range) to the destination range
            sourceRange.Copy(destRange);

            // Define output file path
            string outputPath = "TableToRangeCopyDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// To execute the demo
class Program
{
    static void Main()
    {
        TableToRangeCopyDemo.Run();
    }
}