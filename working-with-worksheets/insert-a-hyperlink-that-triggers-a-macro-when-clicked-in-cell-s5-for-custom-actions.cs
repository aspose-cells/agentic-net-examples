// Title: Insert a hyperlink in cell S5 that launches a VBA macro using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add hyperlink to a cell that runs a VBA macro | how to link a macro to a specific cell in an Xlsm workbook with Aspose.Cells | C# create macro‑enabled Excel file and assign hyperlink to MyMacro | Aspose.Cells add hyperlink for custom action on a worksheet cell | save workbook as .xlsm with hyperlink that triggers a macro using Aspose.Cells
// Tags: add hyperlink to cell Aspose.Cells | hyperlink to VBA macro .NET | macro‑enabled workbook creation Aspose.Cells | save as Xlsm with macro link | cell S5 hyperlink macro trigger

using System;
using System.IO;
using Aspose.Cells;

// The example creates a macro‑enabled workbook, writes "Run Macro" into cell S5, adds a hyperlink that points to the VBA macro "MyMacro", ensures the output directory exists, and saves the file as "HyperlinkMacro.xlsm" using the Xlsm format.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new macro‑enabled workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set display text in cell S5
            Cell cell = sheet.Cells["S5"];
            cell.PutValue("Run Macro");

            // Add a hyperlink that points to a macro named "MyMacro"
            // Hyperlink covers cell S5 (row index 4, column index 18)
            // The overload without HyperlinkType defaults to a URL hyperlink
            sheet.Hyperlinks.Add(4, 18, 1, 1, "MyMacro");

            // Define output file path
            string outputPath = "HyperlinkMacro.xlsm";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a macro‑enabled file
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
