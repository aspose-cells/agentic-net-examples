// Title: Log the size of a merged Excel workbook after saving with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells get file size of saved merged workbook | how to measure output Excel file size after combining workbooks using Aspose.Cells | retrieve merged workbook size in kilobytes after SaveFormat.Xlsx in .NET | log saved workbook size when merging multiple .xlsx files with Aspose.Cells | determine storage impact of workbook merge using Aspose.Cells C#
// Tags: Aspose.Cells merged workbook size logging | C# retrieve saved workbook file length | Aspose.Cells SaveFormat.Xlsx size check | remove initial empty worksheet after merge Aspose.Cells | monitor storage impact of workbook combination

using Aspose.Cells;
using System;
using System.IO;

// The example merges several Excel files into a single workbook, removes the automatically created empty sheet, saves the result as an Xlsx file, and writes the final file size in kilobytes to the console for storage impact monitoring.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will hold the merged worksheets
            Workbook mergedWorkbook = new Workbook();

            // Paths of workbooks to be merged
            string[] sourceFiles = { "Book1.xlsx", "Book2.xlsx", "Book3.xlsx" };

            // Load each source workbook and copy its worksheets into the merged workbook
            foreach (string filePath in sourceFiles)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Warning: Source file \"{filePath}\" not found. Skipping.");
                    continue;
                }

                // Load source workbook
                Workbook srcWorkbook = new Workbook(filePath);

                // Copy each worksheet from the source workbook
                foreach (Worksheet ws in srcWorkbook.Worksheets)
                {
                    // AddCopy expects the worksheet name, not the Worksheet object
                    mergedWorkbook.Worksheets.AddCopy(ws.Name);
                }
            }

            // Remove the initially created empty sheet if it still exists
            if (mergedWorkbook.Worksheets.Count > 1)
            {
                Worksheet firstSheet = mergedWorkbook.Worksheets[0];
                bool isEmpty = firstSheet.Cells.MaxDataRow == -1 && firstSheet.Cells.MaxDataColumn == -1;
                if (isEmpty)
                {
                    mergedWorkbook.Worksheets.RemoveAt(0);
                }
            }

            // Save the merged workbook
            string outputPath = "MergedWorkbook.xlsx";
            mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            // Log the file size after saving
            FileInfo fileInfo = new FileInfo(outputPath);
            long sizeInBytes = fileInfo.Length;
            Console.WriteLine($"Merged workbook size: {sizeInBytes / 1024.0:F2} KB");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
