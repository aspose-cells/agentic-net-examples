// Title: C# – Load an Excel 97‑2003 (.xls) workbook with Aspose.Cells using LoadOptions (Excel97To2003)
// Description: Demonstrates how to create a LoadOptions object with LoadFormat.Excel97To2003, verify the .xls file exists, load it via the Workbook constructor, and display the worksheet count.
// Keywords: Aspose.Cells | C# | .NET | LoadOptions | Excel97To2003 | XLS | legacy Excel | open workbook | LoadFormat enumeration | sample code | GitHub example
// Common Searches: open .xls file Aspose.Cells C# | Aspose.Cells LoadOptions Excel97To2003 example | load legacy Excel workbook .NET | Workbook constructor with LoadOptions | check file exists before loading Aspose.Cells
// Developer Intent: Load a legacy .xls workbook by explicitly setting the Excel97To2003 format in LoadOptions.
// Use Cases: Validate the presence of an .xls file to prevent FileNotFoundException. | Force Aspose.Cells to interpret a file as Excel 97‑2003 regardless of its extension. | Retrieve basic workbook information, such as the number of worksheets, after loading.
// AI Prompts: Show C# code that opens an .xls file with Aspose.Cells using LoadOptions set to Excel97To2003 and includes file‑not‑found handling. | Provide an example that loads a legacy Excel workbook and lists all worksheet names with Aspose.Cells. | Explain how the LoadFormat enumeration can be used to enforce a specific Excel version when loading a workbook in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Demonstrates how to create a LoadOptions object with LoadFormat.Excel97To2003, verify the .xls file exists, load it via the Workbook constructor, and display the worksheet count.
    class Program
    {
        static void Main()
        {
            // Path to the Excel 97‑2003 workbook (XLS file)
            string filePath = "sample.xls";

            try
            {
                // Verify the file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Create LoadOptions with the desired format (Excel 97‑2003)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

                // Load the workbook using the file path and the configured LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Demonstrate that the workbook is loaded (e.g., print worksheet count)
                Console.WriteLine("Number of worksheets loaded: " + workbook.Worksheets.Count);
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
