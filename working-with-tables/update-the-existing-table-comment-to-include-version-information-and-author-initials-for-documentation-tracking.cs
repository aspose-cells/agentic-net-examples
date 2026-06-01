using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class UpdateTableCommentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (A1:B3)
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue(200);

                // Add a ListObject (table) covering the data range
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int listObjectIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject listObject = worksheet.ListObjects[listObjectIndex];

                // Existing comment (if any) can be retrieved; we now append version info and author initials
                string existingComment = listObject.Comment ?? string.Empty;
                string versionInfo = "v1.2";
                string authorInitials = "AB";

                // Build the new comment string
                // Example format: "Original comment. Version: v1.2, Author: AB"
                string newComment = string.IsNullOrWhiteSpace(existingComment)
                    ? $"Version: {versionInfo}, Author: {authorInitials}"
                    : $"{existingComment} Version: {versionInfo}, Author: {authorInitials}";

                // Update the table comment
                listObject.Comment = newComment;

                // Output the updated comment to console for verification
                Console.WriteLine("Updated Table Comment: " + listObject.Comment);

                // Save the workbook
                workbook.Save("UpdatedTableComment.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while updating the table comment: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                UpdateTableCommentDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}