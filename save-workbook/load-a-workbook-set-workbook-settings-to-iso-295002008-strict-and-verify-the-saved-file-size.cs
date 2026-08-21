// Title: C# – Load Excel workbook, apply ISO 29500‑2008 Strict compliance, save and verify file size with Aspose.Cells
// Description: Demonstrates how to load an existing .xlsx file using Aspose.Cells for .NET, set the workbook's OOXML compliance to ISO 29500‑2008 Strict via Settings.Compliance, save the workbook, and retrieve the saved file size with FileInfo. Includes an optional reload to confirm the compliance flag persists.
// Keywords: Aspose.Cells C# ISO 29500 strict | set OoxmlCompliance Iso29500_2008_Strict | save workbook and get file size | verify Excel compliance after save | Aspose.Cells file size check | C# load and save Excel strict mode
// Common Searches: Aspose.Cells set ISO 29500 strict compliance C# | how to get saved Excel file size with Aspose.Cells | verify compliance setting after saving workbook Aspose | C# load workbook change OOXML compliance | Aspose.Cells OoxmlCompliance example
// Developer Intent: Set strict OOXML compliance on a workbook, save it, and confirm both the file size and that the compliance setting remains after reload.
// Use Cases: Produce Excel files that must meet ISO 29500‑2008 Strict standards for regulatory compliance. | Measure storage impact of strict compliance versus default mode. | Automated validation pipelines that ensure the compliance flag is retained after file generation.
// AI Prompts: Generate C# code with Aspose.Cells to load a .xlsx, set OoxmlCompliance.Iso29500_2008_Strict, save to a new path, and output the saved file size. | Create a reusable C# method that accepts input and output paths, applies ISO 29500‑2008 strict compliance, saves the workbook, and returns the file size and compliance value. | Explain how to programmatically confirm that the strict compliance setting persists after reloading the saved workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to load an existing .xlsx file using Aspose.Cells for .NET, set the workbook's OOXML compliance to ISO 29500‑2008 Strict via Settings.Compliance, save the workbook, and retrieve the saved file size with FileInfo. Includes an optional reload to confirm the compliance flag persists.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with an actual file path)
            string sourcePath = "source.xlsx";

            // Load the existing workbook using the provided constructor rule
            Workbook workbook = new Workbook(sourcePath);

            // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Define the output file path
            string outputPath = "strict_compliance.xlsx";

            // Save the workbook using the provided Save(string) rule
            workbook.Save(outputPath);

            // Verify the saved file size
            FileInfo fileInfo = new FileInfo(outputPath);
            Console.WriteLine($"Saved file size: {fileInfo.Length} bytes");

            // Optional: reload the saved workbook to confirm the compliance setting persisted
            Workbook reloaded = new Workbook(outputPath);
            Console.WriteLine($"Reloaded workbook compliance: {reloaded.Settings.Compliance}");
        }
    }
}
