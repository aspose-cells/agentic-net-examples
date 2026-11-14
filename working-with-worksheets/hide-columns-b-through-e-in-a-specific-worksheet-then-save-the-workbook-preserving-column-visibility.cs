// Title: Hide columns B‑E in an Aspose.Cells worksheet and save the workbook with hidden columns preserved (C#)
// Common Searches: Aspose.Cells C# hide multiple columns in a worksheet | how to keep columns hidden when saving Excel file with Aspose.Cells | C# code to hide columns B to E using Aspose.Cells API | save workbook as XLSX while preserving column visibility Aspose.Cells | set column IsHidden property for a range of columns in Aspose.Cells .NET
// Tags: Aspose.Cells column IsHidden API | hide column range in worksheet C# | preserve hidden columns on Excel save Aspose | ensure output directory exists before workbook save | save workbook to Xlsx format Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates (or loads) a workbook, accesses the first worksheet, hides columns B through E by setting each column's IsHidden property to true, ensures the target output folder exists, and then saves the file as HiddenColumns.xlsx, preserving the hidden column states. Comprehensive error handling is included.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide columns B (index 1) through E (index 4)
            for (int colIndex = 1; colIndex <= 4; colIndex++)
            {
                // Use the Columns collection to hide a column
                worksheet.Cells.Columns[colIndex].IsHidden = true;
            }

            // Define output file path
            string outputPath = "HiddenColumns.xlsx";

            // Ensure the directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook; column visibility is preserved in the saved file
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
