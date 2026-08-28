// Title: Automatically create a second worksheet for CSV rows that exceed Excel's row limit using Aspose.Cells TxtLoadOptions in C#
// AI Prompts: Generate C# code that loads a CSV with Aspose.Cells and uses TxtLoadOptions.ExtendToNextSheet to automatically spill rows beyond the Excel row limit into a new worksheet. | Show how to retrieve the worksheet count and row counts after loading a large CSV with auto‑populate enabled in Aspose.Cells. | Write an example that saves the workbook after overflow rows have been moved to a secondary sheet using TxtLoadOptions.
// Common Searches: Aspose.Cells C# load CSV and split into multiple worksheets when row count exceeds Excel limit | How to use TxtLoadOptions ExtendToNextSheet property in .NET | C# example for handling Excel maximum rows with Aspose.Cells | Auto‑populate overflow rows to next sheet using Aspose.Cells TxtLoadOptions | Saving overflow CSV data to separate worksheets with Aspose.Cells
// Tags: Aspose.Cells TxtLoadOptions ExtendToNextSheet | auto‑populate overflow rows to new worksheet | load CSV into workbook with sheet overflow | C# split large CSV across Excel worksheets | handling Excel row limit with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates using TxtLoadOptions.ExtendToNextSheet to load a CSV that exceeds Excel's maximum row count, automatically creating a second worksheet for overflow rows and saving the result as an XLSX file.
class AutoPopulateExample
{
    static void Main()
    {
        // Generate CSV data that exceeds Excel's maximum row limit (1,048,576 rows)
        int totalRows = 1_048_580; // 4 rows will overflow to the next sheet
        using (MemoryStream csvStream = new MemoryStream())
        using (StreamWriter writer = new StreamWriter(csvStream))
        {
            // Write header
            writer.WriteLine("ID,Value");

            // Write data rows
            for (int i = 1; i <= totalRows; i++)
            {
                writer.WriteLine($"{i},Data{i}");
            }

            writer.Flush();
            csvStream.Position = 0; // Reset stream position for reading

            // Load the CSV into a workbook with auto‑populate enabled
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                ExtendToNextSheet = true // Spill excess rows to a new worksheet
            };

            // Create workbook from the CSV stream using the load options
            Workbook workbook = new Workbook(csvStream, loadOptions);

            // Output information about the resulting workbook
            Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
            Console.WriteLine("Rows in first worksheet: " + (workbook.Worksheets[0].Cells.MaxDataRow + 1));
            Console.WriteLine("Rows in second worksheet: " + (workbook.Worksheets[1].Cells.MaxDataRow + 1));

            // Save the workbook to a file
            workbook.Save("AutoPopulated.xlsx");
        }
    }
}
