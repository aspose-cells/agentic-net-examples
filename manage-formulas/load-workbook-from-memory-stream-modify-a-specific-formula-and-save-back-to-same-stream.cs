// Title: How to update a cell formula in an Excel workbook stored in a MemoryStream using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change formula of a specific cell in a workbook loaded from MemoryStream | update Excel cell formula without writing to a temporary file using Aspose.Cells | replace worksheet cell formula and rewrite the same MemoryStream with Aspose.Cells | reset MemoryStream position before saving modified workbook in Aspose.Cells | C# load workbook from MemoryStream, modify a formula, and save back to the stream
// Tags: update cell formula Aspose.Cells | load workbook from MemoryStream Aspose.Cells | save workbook to MemoryStream SaveFormat.Xlsx | reset stream position before saving Aspose.Cells | modify worksheet cell formula C#

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookFormulaUpdaterApp
{
    // The UpdateFormula method positions the input MemoryStream at the start, loads it into an Aspose.Cells Workbook, accesses the specified worksheet and cell, assigns a new formula, clears the stream, saves the workbook back to the same stream in XLSX format, and finally resets the stream position for subsequent reading.
    public class WorkbookFormulaUpdater
    {
        /// <param name="stream">MemoryStream containing the workbook data.</param>
        /// <param name="sheetIndex">Zero‑based index of the worksheet to modify.</param>
        /// <param name="cellName">A1‑style address of the cell whose formula will be changed.</param>
        /// <param name="newFormula">The new formula string (e.g., "=SUM(B1:B10)").</param>
        public void UpdateFormula(MemoryStream stream, int sheetIndex, string cellName, string newFormula)
        {
            try
            {
                // Ensure the stream is positioned at the beginning before loading.
                stream.Position = 0;

                // Load the workbook from the memory stream.
                Workbook workbook = new Workbook(stream);

                // Access the target worksheet.
                Worksheet worksheet = workbook.Worksheets[sheetIndex];

                // Update the formula of the specified cell.
                Cell targetCell = worksheet.Cells[cellName];
                targetCell.Formula = newFormula;

                // Prepare the stream for writing: reset position and truncate existing content.
                stream.Position = 0;
                stream.SetLength(0);

                // Save the modified workbook back into the same memory stream.
                workbook.Save(stream, SaveFormat.Xlsx);

                // Reset the position so the caller can read from the beginning.
                stream.Position = 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating formula: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a sample workbook with some data.
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].PutValue(30);

                // Save the workbook to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Save(ms, SaveFormat.Xlsx);

                    // Update the formula in cell B1.
                    WorkbookFormulaUpdater updater = new WorkbookFormulaUpdater();
                    updater.UpdateFormula(ms, 0, "B1", "=SUM(A1:A3)");

                    // Write the modified workbook to a file.
                    string outputPath = "UpdatedWorkbook.xlsx";
                    File.WriteAllBytes(outputPath, ms.ToArray());
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
