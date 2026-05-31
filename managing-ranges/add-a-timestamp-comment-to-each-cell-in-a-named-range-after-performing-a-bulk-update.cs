using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTimestampComments
{
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

                // Populate sample data (A1:C3)
                for (int r = 0; r < 3; r++)
                {
                    for (int c = 0; c < 3; c++)
                    {
                        cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                    }
                }

                // Define a named range that covers A1:B2
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                namedRange.RefersTo = "=Sheet1!$A$1:$B$2";

                // Perform a bulk update on the named range (set all cells to "Updated")
                string address = namedRange.RefersTo.TrimStart('='); // remove leading '='
                AsposeRange range = cells.CreateRange(address);
                // Use overload that accepts conversion flags (false, false) to set a plain string value
                range.PutValue("Updated", false, false);

                // Add a timestamp comment to each cell in the named range
                CommentCollection comments = sheet.Comments;
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
                {
                    for (int c = range.FirstColumn; c < range.FirstColumn + range.ColumnCount; c++)
                    {
                        int commentIdx = comments.Add(r, c);
                        Comment comment = comments[commentIdx];
                        comment.Note = $"Updated on {timestamp}";
                    }
                }

                // Ensure output directory exists (handle case where outputPath has no directory part)
                string outputPath = "TimestampComments.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}