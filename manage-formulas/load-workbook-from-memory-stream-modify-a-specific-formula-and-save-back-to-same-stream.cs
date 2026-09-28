// Title: Update a cell formula in a MemoryStream‑based Excel workbook and rewrite it to the same stream with Aspose.Cells for .NET
// AI Prompts: Load a workbook from a MemoryStream, assign a new formula to a given cell, and write the updated workbook back into the same stream using Aspose.Cells for .NET. | Overwrite the original MemoryStream after changing the formula of cell A1 on sheet "Sheet1" with Aspose.Cells, ensuring the stream is truncated and repositioned. | Ensure a worksheet exists, set its cell formula, and save the workbook in‑place to the same MemoryStream without creating a temporary file.
// Common Searches: Aspose.Cells C# modify formula in Excel file stored in a MemoryStream | How to save changes back to the same MemoryStream after editing an Excel workbook with Aspose.Cells | Replace Excel cell formula in memory and keep the stream open using Aspose.Cells | C# update worksheet formula without writing to disk with Aspose.Cells | Overwrite original MemoryStream after changing Excel formulas in .NET
// Tags: cell formula update Aspose.Cells | in‑place workbook stream save | memory‑based Excel modification C# | stream truncation before save Aspose | worksheet creation if missing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an XLSX workbook from a MemoryStream with Aspose.Cells, ensuring the target worksheet exists, changing the formula of a specified cell, truncating the original stream, saving the workbook back into the same stream, and resetting the stream position for further processing.
public class WorkbookProcessor
{
    /// <param name="stream">MemoryStream containing the original workbook data.</param>
    /// <param name="sheetName">Name of the worksheet that contains the target cell.</param>
    /// <param name="cellName">A1‑style address of the cell whose formula will be changed.</param>
    /// <param name="newFormula">The new formula string (e.g., "=SUM(B1:B10)").</param>
    public void UpdateFormulaInStream(MemoryStream stream, string sheetName, string cellName, string newFormula)
    {
        try
        {
            // Ensure the stream is positioned at the beginning before loading.
            stream.Position = 0;

            // Load the workbook from the memory stream.
            Workbook workbook = new Workbook(stream);

            // Access the required worksheet; create it if it does not exist.
            Worksheet worksheet = workbook.Worksheets[sheetName] ?? workbook.Worksheets.Add(sheetName);

            // Access the target cell and set the new formula.
            Cell targetCell = worksheet.Cells[cellName];
            targetCell.Formula = newFormula;

            // Prepare the stream for writing the updated workbook.
            stream.SetLength(0); // Truncate the existing content.
            workbook.Save(stream, SaveFormat.Xlsx); // Save back into the same stream.
            stream.Position = 0; // Reset position for the caller.
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"UpdateFormulaInStream error: {ex.Message}");
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
            var processor = new WorkbookProcessor();

            // Create a new workbook with a default worksheet.
            Workbook wb = new Workbook();
            wb.Worksheets[0].Name = "Sheet1";

            // Save the workbook to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Save(ms, SaveFormat.Xlsx);

                // Update a formula in the workbook.
                processor.UpdateFormulaInStream(ms, "Sheet1", "A1", "=SUM(B1:B10)");

                // Write the updated workbook to a file for verification.
                string outputPath = "UpdatedWorkbook.xlsx";
                File.WriteAllBytes(outputPath, ms.ToArray());
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
