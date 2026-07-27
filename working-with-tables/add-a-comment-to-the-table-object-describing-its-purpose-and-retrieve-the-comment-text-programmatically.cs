using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCommentDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the table (list object)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Add a ListObject (table) covering the data range A1:B3, with header row
            int listObjectIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject listObj = worksheet.ListObjects[listObjectIndex];

            // Set a comment describing the purpose of the table
            listObj.Comment = "Employee data table: stores ID and Name of staff members";

            // Retrieve the comment text programmatically
            string retrievedComment = listObj.Comment;
            Console.WriteLine("Table Comment: " + retrievedComment);

            // Save the workbook
            workbook.Save("TableCommentDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}