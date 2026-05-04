using System;
using System.IO;
using Aspose.Cells;

class ThreadedCommentExample
{
    static void Main()
    {
        // Load an existing workbook into memory (replace with your own source if needed)
        // Here we read a file into a byte array and then create a MemoryStream.
        byte[] workbookBytes = File.ReadAllBytes("Template.xlsx");
        using (MemoryStream inputStream = new MemoryStream(workbookBytes))
        {
            // Load the workbook from the memory stream
            Workbook workbook = new Workbook(inputStream);

            // Access the first worksheet (you can choose any worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Add a threaded comment author (if the author already exists you can retrieve it by index)
            int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
                "John Doe",               // Author name
                "john.doe@example.com",  // User ID / email
                "EXAMPLE_PROVIDER");     // Provider (can be empty)
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

            // Insert a threaded comment into cell B2 using the cell name overload
            sheet.Comments.AddThreadedComment("B2", "This is a threaded comment.", author);

            // Retrieve the threaded comments for verification (optional)
            ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments("B2");
            foreach (ThreadedComment tc in threadedComments)
            {
                Console.WriteLine($"Author: {tc.Author.Name}, Text: {tc.Notes}");
            }

            // Save the modified workbook back to a memory stream
            using (MemoryStream outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);

                // For demonstration, write the updated workbook to a file
                File.WriteAllBytes("UpdatedWorkbook.xlsx", outputStream.ToArray());
            }
        }
    }
}