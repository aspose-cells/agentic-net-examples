// Title: Save an Excel workbook that contains slicers to Excel 2016 (XLSX) format while preserving slicer functionality using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with slicers using Aspose.Cells, then save it as an Excel 2016 XLSX file with OoxmlSaveOptions to keep the slicer objects intact. | After optionally modifying slicer properties in a worksheet, export the workbook with OoxmlSaveOptions so that all slicer features remain functional in the saved file.
// Common Searches: how to keep slicers when converting an Excel file with Aspose.Cells C# | Aspose.Cells save workbook with slicers to Excel 2016 format | C# OoxmlSaveOptions preserve slicer objects in XLSX | export slicer-enabled worksheet using Aspose.Cells .NET | retain slicer functionality after saving Excel file with Aspose.Cells
// Tags: save workbook with slicers using OoxmlSaveOptions | Aspose.Cells slicer preservation in XLSX | export slicer-enabled Excel file C# | Excel 2016 format retention of slicer objects | load and save workbook containing slicers Aspose

using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for slicer support
using Aspose.Cells.Saving;    // Required for OoxmlSaveOptions
using System;
using System.IO;

// The example checks for the input file, loads the workbook that includes slicers with Aspose.Cells, optionally accesses or modifies slicer properties, configures OoxmlSaveOptions for the XLSX format, and saves the workbook as output.xlsx, ensuring that all slicer objects are retained in the Excel 2016 file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains slicers
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Manipulate slicers here if needed
            // Example: access the first slicer in the first worksheet
            // Worksheet sheet = workbook.Worksheets[0];
            // SlicerCollection slicers = sheet.Slicers;
            // if (slicers.Count > 0)
            // {
            //     Slicer slicer = slicers[0];
            //     slicer.Name = "MySlicer";
            // }

            // Prepare save options for Excel 2016 (XLSX) format
            // OoxmlSaveOptions does not expose a Version property in all library versions,
            // so we rely on the default behavior which preserves advanced features like slicers.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook; slicers will be retained in the output file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
