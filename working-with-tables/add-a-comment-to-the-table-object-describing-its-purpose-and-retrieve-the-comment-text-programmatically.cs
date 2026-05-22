using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ListObjectCommentDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the table
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue(200);

        // Add a ListObject (table) that covers the data range A1:B3
        int listIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = worksheet.ListObjects[listIndex];

        // Set a comment describing the purpose of the table
        table.Comment = "This table stores IDs and their corresponding values.";

        // Retrieve the comment text programmatically and display it
        Console.WriteLine("Table Comment: " + table.Comment);

        // Save the workbook
        workbook.Save("ListObjectWithComment.xlsx", SaveFormat.Xlsx);
    }
}