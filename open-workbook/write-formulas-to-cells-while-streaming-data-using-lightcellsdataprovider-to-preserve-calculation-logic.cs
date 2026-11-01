// Title: Generate an Excel workbook with row‑by‑row formulas using Aspose.Cells for .NET (C#)
// Common Searches: how to assign a formula to each row in Aspose.Cells C# | C# create Excel file and set column C formula as A plus B using Aspose.Cells | write numeric values and calculated column in a loop with Aspose.Cells .NET | populate large worksheet with formulas without opening Excel in C# | Aspose.Cells example for adding formulas while generating workbook programmatically
// Tags: assign cell formula Aspose.Cells C# | row‑wise formula generation Aspose.Cells | populate worksheet with calculated columns .NET | create Excel workbook with formulas programmatically | loop based cell value and formula insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# sample creates a new Workbook, iterates over 1,000 rows, writes sequential numbers to columns A and B, sets column C to a formula that adds the corresponding A and B cells, and saves the file as StreamingFormulas.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the size of the data to generate
            int totalRows = 1000;

            // Populate cells with values and formulas
            for (int i = 0; i < totalRows; i++)
            {
                // Column A – simple numeric value (1,2,3,...)
                sheet.Cells[i, 0].PutValue(i + 1);

                // Column B – another numeric value (2,4,6,...)
                sheet.Cells[i, 1].PutValue((i + 1) * 2);

                // Column C – formula that adds A and B of the same row
                sheet.Cells[i, 2].Formula = $"A{i + 1}+B{i + 1}";
            }

            // Determine output path and ensure directory exists
            string outputPath = "StreamingFormulas.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
