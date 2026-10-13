// Title: Insert a column at index 5 and shift existing columns right in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Add a column at zero‑based position 5, shift all later columns right, and copy the style from the neighboring column using Aspose.Cells in C#. | Programmatically insert a new worksheet column at index five with style inheritance via the Aspose.Cells .NET library.
// Common Searches: How to insert a column at index 5 in an Excel file using Aspose.Cells for .NET | Aspose.Cells C# shift columns right after inserting a new column | Copy style from adjacent column when adding a column with Aspose.Cells API
// Tags: column insertion by index Aspose.Cells C# | shift columns right Aspose.Cells | copy adjacent column style Aspose.Cells | zero‑based column insertion .NET | add worksheet column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a new workbook, accesses the first worksheet, inserts a column at zero‑based index 5 while copying the adjacent column's style, shifts existing columns to the right, saves the file as Result.xlsx, and prints the full output path.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Insert one column at index 5 (zero‑based), shifting existing columns to the right
                // The second parameter indicates whether to copy the style from the adjacent column
                worksheet.Cells.InsertColumn(5, true);

                // Define the output file path
                string outputPath = "Result.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
