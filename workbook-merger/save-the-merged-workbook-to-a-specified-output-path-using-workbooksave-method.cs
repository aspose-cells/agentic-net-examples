// Title: Save a Merged Workbook to a Specific Path with Aspose.Cells C#
// Description: This example creates a source and a destination workbook, merges them using Workbook.Combine, and then persists the combined workbook to "CombinedWorkbook.xlsx" with the Workbook.Save(string, SaveFormat.Xlsx) overload.
// Keywords: Aspose.Cells | C# | Workbook.Combine | Workbook.Save | save merged workbook | output file path | XLSX format | combine Excel files | .NET Excel export
// Common Searches: Aspose.Cells save combined workbook C# | How to use Workbook.Combine and Save in .NET | C# merge two Excel files with Aspose.Cells | Save merged workbook to specific folder Aspose.Cells | Workbook.Save overload filename format
// Developer Intent: Persist the result of a Workbook.Combine operation to a designated file location.
// Use Cases: Merge a template workbook with a data workbook and generate a final report file. | Consolidate departmental spreadsheets into a single archive workbook on a shared drive. | Automate creation of a combined financial statement and store it in a cloud folder. | Generate a master workbook from multiple source files for batch processing. | Publish a merged workbook to a web server after saving it locally.
// AI Prompts: Write C# code that merges three Excel workbooks using Aspose.Cells and saves the result to a network share with SaveFormat.Xlsx. | Show how to handle exceptions when saving a merged workbook to a read‑only directory with Aspose.Cells. | Demonstrate saving a combined workbook in different formats (XLSX, CSV, PDF) after using Workbook.Combine.

using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndSaveDemo
{
    // This example creates a source and a destination workbook, merges them using Workbook.Combine, and then persists the combined workbook to "CombinedWorkbook.xlsx" with the Workbook.Save(string, SaveFormat.Xlsx) overload.
    class Program
    {
        static void Main()
        {
            // Create the source workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Source";
            sourceSheet.Cells["A1"].PutValue("Source Data");

            // Create the destination workbook and add sample data
            Workbook destWorkbook = new Workbook(FileFormatType.Xlsx);
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Destination";
            destSheet.Cells["B2"].PutValue("Destination Data");

            // Merge the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Define the output file path
            string outputPath = "CombinedWorkbook.xlsx";

            // Save the merged workbook using the Save(string, SaveFormat) overload
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Merged workbook saved to: {outputPath}");
        }
    }
}
