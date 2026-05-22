using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class TableToRangeSortDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate header row
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Score");
                cells["D1"].PutValue("Category");

                // Populate sample data rows
                cells["A2"].PutValue(1); cells["B2"].PutValue("Alice");   cells["C2"].PutValue(85); cells["D2"].PutValue("X");
                cells["A3"].PutValue(2); cells["B3"].PutValue("Bob");     cells["C3"].PutValue(92); cells["D3"].PutValue("Y");
                cells["A4"].PutValue(3); cells["B4"].PutValue("Charlie"); cells["C4"].PutValue(78); cells["D4"].PutValue("X");
                cells["A5"].PutValue(4); cells["B5"].PutValue("Diana");   cells["C5"].PutValue(88); cells["D5"].PutValue("Y");

                // Add a table (ListObject) covering the data range A1:D5
                int tableIndex = sheet.ListObjects.Add("A1", "D5", true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SampleTable";

                // Capture the data range of the table (including header)
                AsposeRange tableRange = table.DataRange; // includes header row

                // Convert the table to a normal range (the table object will be removed)
                table.ConvertToRange();

                // Verify that the table has been removed
                Console.WriteLine("ListObjects count after conversion: " + sheet.ListObjects.Count);

                // Set up the DataSorter to sort by the "Score" column (third column, index 2)
                DataSorter sorter = workbook.DataSorter;
                sorter.HasHeaders = true;               // First row is header
                sorter.AddKey(2, SortOrder.Descending); // Sort by column C (Score) descending

                // Define the sort area using the previously captured range address
                CellArea sortArea = new CellArea
                {
                    StartRow = tableRange.FirstRow,
                    StartColumn = tableRange.FirstColumn,
                    EndRow = tableRange.FirstRow + tableRange.RowCount - 1,
                    EndColumn = tableRange.FirstColumn + tableRange.ColumnCount - 1
                };

                // Perform the sort
                sorter.Sort(cells, sortArea);

                // Confirm that sorting affected only the defined range
                Console.WriteLine("Value in cell E1 (should be empty): '" + cells["E1"].StringValue + "'");

                // Output the sorted data to console for verification
                Console.WriteLine("Sorted data (including header):");
                for (int r = sortArea.StartRow; r <= sortArea.EndRow; r++)
                {
                    for (int c = sortArea.StartColumn; c <= sortArea.EndColumn; c++)
                    {
                        Console.Write(cells[r, c].StringValue + "\t");
                    }
                    Console.WriteLine();
                }

                // Save the workbook (overwrite if it already exists)
                string outputPath = "TableToRangeSorted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableToRangeSortDemo.Run();
        }
    }
}