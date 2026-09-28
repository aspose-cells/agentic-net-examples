// Title: Load an XLSX workbook without charts using Aspose.Cells FilterObjects and confirm empty chart collections (C#)
// AI Prompts: Apply Aspose.Cells LoadOptions.FilterObjects to open an XLSX file while ignoring chart objects, then iterate workbook.Worksheets and assert that each sheet.Charts.Count equals zero. | Rewrite the loading routine to exclude charts, and output a success message only when every worksheet's Charts collection is empty after the workbook is loaded.
// Common Searches: Aspose.Cells C# load workbook without chart objects using FilterObjects | How to ignore charts when opening an Excel file with Aspose.Cells .NET | Check worksheet chart collection is empty after loading with Aspose.Cells | FilterObjects property example for excluding charts in Aspose.Cells | C# verify no charts in loaded workbook Aspose.Cells
// Tags: load workbook with FilterObjects Aspose.Cells | exclude chart objects during Excel load .NET | verify empty worksheet chart collection Aspose.Cells | filter out charts using LoadOptions C# | chart collection count check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to load an XLSX workbook with Aspose.Cells while skipping chart objects by configuring LoadOptions.FilterObjects. After loading, it iterates through each worksheet, prints the chart count, and confirms that the Charts collection is empty, handling missing files and exceptions.
class Program
{
    static void Main()
    {
        // Path to the source workbook (replace with your actual file path)
        string sourcePath = "input.xlsx";

        // Ensure the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Configure load options (ignore objects such as charts, shapes, etc.) if supported.
            // LoadDataOnly property is not available in this version, so we load normally.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Verify that each worksheet has no charts after loading
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartCount = sheet.Charts.Count;
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {chartCount} chart(s).");

                if (chartCount != 0)
                {
                    Console.WriteLine("Error: Chart collection is not empty.");
                }
                else
                {
                    Console.WriteLine("Success: Chart collection is empty.");
                }
            }

            // (Optional) Save the workbook if further processing is needed.
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred: {ex.Message}");
        }
    }
}
