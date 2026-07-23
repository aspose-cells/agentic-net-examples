// Title: Assign DocumentVersion = "2.0" to an Excel workbook with Aspose.Cells for .NET
// Description: Loads an existing .xlsx file using Aspose.Cells, sets the built‑in DocumentVersion property to "2.0", and saves the workbook. Shows how to modify Excel metadata programmatically in C#.
// Keywords: Aspose.Cells DocumentVersion | C# set Excel built‑in property | modify workbook metadata | Excel file version tag | Aspose.Cells example .NET
// Common Searches: Aspose.Cells set DocumentVersion C# | change Excel built‑in property programmatically | how to update workbook version metadata with Aspose | C# code to set DocumentVersion in .xlsx | Aspose.Cells modify document properties
// Developer Intent: Update the built‑in DocumentVersion property of a loaded workbook to "2.0" and persist the change.
// Use Cases: Tag reports with a release version before distribution. | Automate version metadata for a batch of generated Excel files. | Integrate version information into a document management workflow.
// AI Prompts: Write C# code that loads an .xlsx with Aspose.Cells, sets DocumentVersion to 2.0, and saves the file. | Explain how to read the DocumentVersion property after it has been set using Aspose.Cells. | Suggest robust error handling for loading a workbook and updating its DocumentVersion when the source file may be missing or corrupted.

using System;
using Aspose.Cells;

namespace AsposeCellsDocumentVersionDemo
{
    // Loads an existing .xlsx file using Aspose.Cells, sets the built‑in DocumentVersion property to "2.0", and saves the workbook. Shows how to modify Excel metadata programmatically in C#.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook to be loaded
            string inputPath = "input.xlsx";

            // Load the workbook from the file (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Set the built‑in DocumentVersion property to "2.0"
            workbook.BuiltInDocumentProperties.DocumentVersion = "2.0";

            // Save the modified workbook (uses the provided save rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"DocumentVersion set to \"2.0\" and workbook saved to {outputPath}");
        }
    }
}
