// Title: Apply a custom string sort order to an Excel column using Aspose.Cells DataSorter in C#
// AI Prompts: Sort column A of a worksheet by the custom list "Critical,High,Medium,Low" with Aspose.Cells DataSorter while keeping the header row intact. | Generate a C# program that defines a custom sort sequence, adds it as a sort key for a specified range, and saves the sorted workbook.
// Common Searches: aspocells c# custom sort order for text column | how to use DataSorter with a predefined list in Aspose.Cells | sorting Excel rows by custom priority list using Aspose.Cells C# | preserve first row as header when sorting with Aspose.Cells DataSorter
// Tags: custom string sort order Aspose.Cells | DataSorter custom order key C# | sort Excel column by predefined sequence Aspose.Cells | header row handling in Aspose.Cells sort C#

using System;
using Aspose.Cells;

namespace CustomSortExample
{
    // The example creates a workbook, fills columns A and B with category and value data, defines a custom order "Critical,High,Medium,Low", configures DataSorter to treat the first row as headers, adds a sort key for column A using the custom order, sorts the range A1:B5, prints the sorted rows, and saves the result as CustomSortedOutput.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data with a header row
                // Column A contains categories that we want to sort using a custom order
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Medium");
                cells["B2"].PutValue(20);
                cells["A3"].PutValue("Low");
                cells["B3"].PutValue(10);
                cells["A4"].PutValue("High");
                cells["B4"].PutValue(30);
                cells["A5"].PutValue("Critical");
                cells["B5"].PutValue(40);

                // Define the custom sort list (desired sequence)
                string customOrder = "Critical,High,Medium,Low";

                // Configure the DataSorter
                DataSorter sorter = workbook.DataSorter;
                sorter.HasHeaders = true; // First row is a header and should stay at the top
                // Add a sort key for column A (index 0) with ascending order and the custom list
                sorter.AddKey(0, SortOrder.Ascending, customOrder);

                // Define the range to sort (including header row)
                CellArea sortArea = CellArea.CreateCellArea("A1", "B5");

                // Perform the sort
                sorter.Sort(cells, sortArea);

                // Output the sorted result to the console
                Console.WriteLine("Sorted data using custom order:");
                int lastDataRow = cells.MaxDataRow; // zero‑based index of the last row with data
                for (int row = 1; row <= lastDataRow; row++) // start from row 2 (index 1) to skip header
                {
                    Console.WriteLine($"{cells[row, 0].StringValue} : {cells[row, 1].IntValue}");
                }

                // Save the workbook (optional)
                workbook.Save("CustomSortedOutput.xlsx");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
