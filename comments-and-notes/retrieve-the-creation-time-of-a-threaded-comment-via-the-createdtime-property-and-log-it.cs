// Title: Read the first regular comment's note from an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to get the note of the first comment in an Excel sheet with Aspose.Cells C# | Aspose.Cells retrieve comment text from the first worksheet in a workbook | C# example for checking and reading worksheet comments using Aspose.Cells | load an Excel file and enumerate its comments with Aspose.Cells .NET | handle file-not-found errors when accessing comments in Aspose.Cells
// Tags: Aspose.Cells read comment note C# | Aspose.Cells worksheet comments collection .NET | Aspose.Cells load workbook from file C# | Aspose.Cells verify comment existence before access | Aspose.Cells exception handling for missing Excel file

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that 'sample.xlsx' exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, checks for regular comments, retrieves the first Comment object's Note property, and writes the note to the console while handling possible exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string filePath = "sample.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Check for regular comments in the worksheet (threaded comments are not supported in this version)
            if (sheet.Comments != null && sheet.Comments.Count > 0)
            {
                // Retrieve the first comment
                Comment firstComment = sheet.Comments[0];

                // Output the comment's note text
                Console.WriteLine($"First comment note: {firstComment.Note}");
            }
            else
            {
                Console.WriteLine("No comments found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
