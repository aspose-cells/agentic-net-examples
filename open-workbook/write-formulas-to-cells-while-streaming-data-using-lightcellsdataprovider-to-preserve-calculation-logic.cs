// Title: Add column formulas while bulk‑populating a large Excel worksheet with Aspose.Cells for .NET
// Common Searches: how to assign a formula to each row when inserting thousands of rows with Aspose.Cells C# | write Excel formulas during bulk data import using Aspose.Cells .NET API | calculate all formulas after programmatically filling a large worksheet with Aspose.Cells | C# example for saving a large workbook after adding formulas to a column | workaround LightCellsDataProvider compile issue by writing formulas directly in Aspose.Cells
// Tags: Aspose.Cells bulk insert formulas C# | write column B formula referencing column A Aspose.Cells | calculate workbook after data population Aspose.Cells | save large Excel file with formulas Aspose.Cells .NET | LightCellsDataProvider alternative formula assignment

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamingExample
{
    // The example creates a new Workbook, iterates over 10,000 rows, writes a numeric value to column A and assigns a formula (e.g., =A1*2) to column B for each row, forces formula evaluation with CalculateFormula(), ensures the output directory exists, and saves the result as StreamingFormulas.xlsx, demonstrating a practical workaround for LightCellsDataProvider limitations when adding formulas during large‑scale data population.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the number of rows and columns to populate
                int totalRows = 10000;      // example large dataset
                int totalColumns = 2;       // column A for values, column B for formulas

                // Populate cells directly (avoids LightCellsDataProvider compile issue)
                for (int row = 0; row < totalRows; row++)
                {
                    // Column A – numeric value
                    worksheet.Cells[row, 0].PutValue(row + 1);

                    // Column B – formula referencing column A
                    worksheet.Cells[row, 1].Formula = $"=A{row + 1}*2";
                }

                // After data import, force calculation so that formulas are evaluated
                workbook.CalculateFormula();

                // Prepare output path and ensure directory exists
                string outputPath = "StreamingFormulas.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
