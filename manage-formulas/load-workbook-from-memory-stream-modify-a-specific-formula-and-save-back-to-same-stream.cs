// Title: Update a specific cell's formula in an Excel workbook loaded from a MemoryStream using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change formula of a cell after loading workbook from MemoryStream | how to modify Excel cell formula in memory without saving to disk using Aspose.Cells | replace cell A1 formula with SUM function in a workbook stored in a MemoryStream | C# update worksheet cell formula and rewrite same MemoryStream with Aspose.Cells | set new formula for a cell in an in‑memory Excel file using Aspose.Cells API
// Tags: update cell formula memory stream Aspose.Cells | modify worksheet cell formula C# | save workbook to same MemoryStream Aspose.Cells | load Excel from MemoryStream edit formula | set cell formula programmatically .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook from a MemoryStream, accesses a worksheet by index, changes the formula of a specified A1‑style cell, clears the original stream, saves the workbook back into the same stream in XLSX format, and resets the stream position for further reading.
public class WorkbookFormulaUpdater
{
    /// <param name="stream">MemoryStream containing the workbook data. The stream will be reset to the beginning after saving.</param>
    /// <param name="sheetIndex">Zero‑based index of the worksheet to modify.</param>
    /// <param name="cellName">A1‑style address of the cell whose formula should be changed (e.g., "A1").</param>
    /// <param name="newFormula">The new formula string (e.g., "=SUM(B1:B10)").</param>
    public static void UpdateFormula(MemoryStream stream, int sheetIndex, string cellName, string newFormula)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (string.IsNullOrEmpty(cellName)) throw new ArgumentException("Cell name must be provided.", nameof(cellName));
        if (newFormula == null) throw new ArgumentNullException(nameof(newFormula));

        try
        {
            // Ensure the stream position is at the beginning before loading.
            stream.Position = 0;

            // Load the workbook from the memory stream.
            Workbook workbook = new Workbook(stream);

            // Access the target worksheet.
            Worksheet worksheet = workbook.Worksheets[sheetIndex];

            // Get the target cell and set the new formula.
            Cell targetCell = worksheet.Cells[cellName];
            targetCell.Formula = newFormula;

            // Prepare the stream for saving: clear existing content.
            stream.SetLength(0);
            stream.Position = 0;

            // Save the modified workbook back into the same stream.
            workbook.Save(stream, SaveFormat.Xlsx);

            // Reset the stream position so callers can read from the beginning.
            stream.Position = 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error updating formula: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook into a memory stream.
            using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(inputPath)))
            {
                // Update the formula in the first worksheet, cell A1.
                WorkbookFormulaUpdater.UpdateFormula(ms, 0, "A1", "=SUM(B1:B10)");

                // Write the modified workbook back to disk.
                File.WriteAllBytes(outputPath, ms.ToArray());
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled error: {ex.Message}");
        }
    }
}
