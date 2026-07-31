// Title: Convert Aspose.Cells Table (ListObject) to a Range and Copy to Another Worksheet – C# Example
// Description: Shows how to create a workbook, add a ListObject table, convert it to a normal range with ConvertToRange, define matching source and destination ranges, copy the data using Range.Copy, and save the workbook.
// Keywords: Aspose.Cells ConvertToRange | Aspose.Cells copy range | Aspose.Cells ListObject to range C# | Range.Copy Aspose.Cells | C# Aspose.Cells table to range | duplicate table data Aspose.Cells | Aspose.Cells worksheet copy example
// Common Searches: Aspose.Cells ConvertToRange C# | How to copy a range to another sheet in Aspose.Cells | Convert ListObject to range Aspose.Cells | Copy table data to another worksheet Aspose.Cells | Range.Copy method Aspose.Cells example
// Developer Intent: Transform a ListObject table into a plain range and duplicate that range on a different worksheet.
// Use Cases: Export table contents as plain data for CSV output or calculations that require non‑table structures. | Create a report sheet that mirrors table data without preserving table features like filters or structured references. | Build a template where table data is copied as a range to allow custom styling, merging, or further processing.
// AI Prompts: Write C# code using Aspose.Cells to convert a ListObject to a range and copy it to another worksheet with Range.Copy. | Explain the steps required to change an Aspose.Cells table into a normal range and then duplicate that range on a different sheet. | Provide an example that copies several tables as ranges to separate destination worksheets using Aspose.Cells in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeCellRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a ListObject table, convert it to a normal range with ConvertToRange, define matching source and destination ranges, copy the data using Range.Copy, and save the workbook.
    public class TableToRangeAndCopyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (source)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate sample data for the table (A1:B3)
                sourceSheet.Cells["A1"].PutValue("ID");
                sourceSheet.Cells["B1"].PutValue("Name");
                sourceSheet.Cells["A2"].PutValue(1);
                sourceSheet.Cells["B2"].PutValue("John");
                sourceSheet.Cells["A3"].PutValue(2);
                sourceSheet.Cells["B3"].PutValue("Mary");

                // Add a ListObject (table) covering the data range
                int tableIndex = sourceSheet.ListObjects.Add("A1", "B3", true);
                ListObject table = sourceSheet.ListObjects[tableIndex];

                // Convert the table to a normal range
                table.ConvertToRange();

                // Define the source range that was previously the table
                AsposeCellRange sourceRange = sourceSheet.Cells.CreateRange("A1:B3");

                // Add a new worksheet to receive the copied range
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Define the destination range (same size) starting at A1 in the destination sheet
                AsposeCellRange destRange = destSheet.Cells.CreateRange("A1:B3");

                // Copy the source range to the destination range
                sourceRange.Copy(destRange);

                // Save the workbook
                string outputPath = "TableToRangeCopyDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableToRangeAndCopyDemo.Run();
        }
    }
}
