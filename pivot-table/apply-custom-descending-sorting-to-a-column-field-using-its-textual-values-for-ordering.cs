using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (first row as header)
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("Banana");
                cells["A3"].PutValue("Apple");
                cells["A4"].PutValue("Cherry");
                cells["A5"].PutValue("Date");

                // Get the DataSorter object
                DataSorter sorter = workbook.DataSorter;

                // Indicate that the range contains a header row
                sorter.HasHeaders = true;

                // Add a sort key for column A (index 0) with descending order
                sorter.AddKey(0, SortOrder.Descending);

                // Define the range to be sorted (including header)
                CellArea sortArea = CellArea.CreateCellArea("A1", "A5");

                // Perform the sort
                sorter.Sort(worksheet.Cells, sortArea);

                // Save the workbook
                string outputPath = "SortedByTextDescending.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}