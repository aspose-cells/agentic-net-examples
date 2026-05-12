using System;
using Aspose.Cells;

namespace UpdateCommentTextDirection
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or iterate through all worksheets as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate over all comments in the worksheet
            foreach (Comment comment in worksheet.Comments)
            {
                // Aspose.Cells does not expose a TextDirection property for Comment in this version.
                // As a workaround, prepend a Left‑to‑Right Mark (LRM) Unicode character to enforce LTR direction.
                comment.Note = "\u200E" + comment.Note;
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}