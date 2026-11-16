// Title: Hide all worksheet comments in an Excel file with Aspose.Cells for .NET to comply with privacy policies
// Common Searches: Aspose.Cells C# hide all comments in a workbook before saving | set Excel comment visibility to false using Aspose.Cells .NET | programmatically remove visibility of cell comments across all sheets with Aspose.Cells | ensure Excel comments are not displayed for privacy using Aspose.Cells API | iterate through worksheets and hide comments in C# Aspose.Cells example
// Tags: hide worksheet comments Aspose.Cells | comment IsVisible property .NET | privacy enforce hidden Excel comments | iterate worksheets Aspose.Cells C# | save workbook with hidden comments

using Aspose.Cells;
using System;

// The example loads an Excel workbook, loops through each worksheet, accesses its CommentCollection, sets each comment's IsVisible property to false to keep comments hidden, and then saves the workbook, ensuring comment visibility complies with privacy requirements.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of comments on the current worksheet
            CommentCollection comments = sheet.Comments;

            // Hide each comment to enforce privacy
            foreach (Comment comment in comments)
            {
                comment.IsVisible = false; // Ensure the comment is not visible to users
            }
        }

        // Save the workbook with the updated comment visibility settings
        workbook.Save("output.xlsx");
    }
}
