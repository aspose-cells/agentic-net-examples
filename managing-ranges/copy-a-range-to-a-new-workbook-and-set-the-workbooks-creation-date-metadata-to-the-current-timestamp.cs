// Title: Copy a Cell Range to a New Workbook and Set Creation Date with Aspose.Cells for .NET (C#)
// Description: Loads source.xlsx, copies cells A1:B2 to C3:D4 in a fresh workbook, updates the workbook's BuiltInDocumentProperties.CreatedTime to the current timestamp, and saves the result as output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# copy range | Excel copy cells | Workbook creation date | BuiltInDocumentProperties | metadata timestamp | range.Copy | Aspose.Cells .NET | Excel automation | copy range between workbooks
// Common Searches: Aspose.Cells copy range to another workbook C# | Set workbook CreatedTime property Aspose.Cells | How to copy cells A1:B2 to C3:D4 using Aspose.Cells | Update Excel file metadata with Aspose.Cells | Copy range and preserve formatting Aspose.Cells
// Developer Intent: Transfer a defined block of cells from an existing Excel file to a new workbook and record the operation time by setting the workbook's creation metadata.
// Use Cases: Generate a lightweight report by extracting a specific data block from a master workbook while automatically stamping the file with the current creation time for audit trails. | Create version‑controlled templates where only the required range is copied to a fresh workbook and the timestamp is stored in the document properties. | Automate data migration between Excel files in a CI/CD pipeline, ensuring each output file carries an accurate CreatedTime for downstream processing.
// AI Prompts: Write C# code using Aspose.Cells to copy cells A1:B2 from source.xlsx to C3:D4 in a new workbook and set BuiltInDocumentProperties.CreatedTime to DateTime.UtcNow. | Explain step‑by‑step how to copy a range to a new workbook and update the workbook's creation date without affecting other worksheets in Aspose.Cells. | Suggest best‑practice error handling and resource cleanup when copying ranges and modifying document properties with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRangeCopyExample
{
    // Loads source.xlsx, copies cells A1:B2 to C3:D4 in a fresh workbook, updates the workbook's BuiltInDocumentProperties.CreatedTime to the current timestamp, and saves the result as output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Define the source range to copy (A1:B2 on the first worksheet)
                Aspose.Cells.Range sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:B2");

                // Create a new (empty) workbook for the destination
                Workbook destWorkbook = new Workbook();

                // Define the destination range where the source data will be copied (C3:D4)
                Aspose.Cells.Range destRange = destWorkbook.Worksheets[0].Cells.CreateRange("C3:D4");

                // Copy the source range into the destination range
                destRange.Copy(sourceRange);

                // Set the workbook's creation date metadata to the current timestamp
                destWorkbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

                // Save the new workbook to a file
                destWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
