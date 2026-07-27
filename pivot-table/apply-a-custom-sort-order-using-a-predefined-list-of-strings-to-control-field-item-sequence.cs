using System;
using Aspose.Cells;

namespace CustomSortExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including a header row)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Medium");
            sheet.Cells["A3"].PutValue("Low");
            sheet.Cells["A4"].PutValue("High");
            sheet.Cells["A5"].PutValue("Critical");

            // Define a custom sort order for the "Category" column
            // The order is: Critical, High, Medium, Low
            string customOrder = "Critical,High,Medium,Low";

            // Configure the DataSorter:
            // - Sort column A (index 0)
            // - Use ascending order (order is ignored when a custom list is supplied)
            // - Apply the custom list defined above
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                     // First row is a header
            sorter.AddKey(0, SortOrder.Ascending, customOrder);

            // Perform the sort on the range that contains the data (including header)
            sorter.Sort(sheet.Cells, CellArea.CreateCellArea("A1", "A5"));

            // Save the result to a file
            workbook.Save("CustomSortedOutput.xlsx");
        }
    }
}