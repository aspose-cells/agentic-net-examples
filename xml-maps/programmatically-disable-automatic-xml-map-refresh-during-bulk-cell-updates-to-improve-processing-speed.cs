// Title: Disable XML Map Refresh for Faster Bulk Cell Updates in Aspose.Cells for .NET
// Description: This C# example loads an existing workbook, verifies the source file, writes values to 10,000 rows × 10 columns, and saves the result. It highlights that the EnableXMLMapRefresh property is not present in the current Aspose.Cells release, so the updates run with the default refresh behavior.
// Keywords: Aspose.Cells .NET | XML map refresh | disable XML map refresh | bulk cell write performance | speed up workbook updates | C# Excel XML mapping | large‑scale cell insertion | Aspose.Cells performance tip
// Common Searches: how to turn off XML map refresh in Aspose.Cells C# | Aspose.Cells bulk cell write speed with XML maps | EnableXMLMapRefresh property missing Aspose.Cells | optimize Excel workbook updates using Aspose.Cells | performance tricks for large data import in Aspose.Cells
// Developer Intent: Prevent automatic XML map refresh while performing massive cell writes to reduce processing time.
// Use Cases: Importing millions of data points into an Excel workbook that contains XML maps. | Running data‑intensive transformations where refresh overhead slows down the job. | Ensuring file‑existence checks before loading to avoid runtime errors. | Saving the workbook after bulk modifications with graceful exception handling.
// AI Prompts: Provide C# code that disables XML map refresh during bulk cell updates with Aspose.Cells, or suggest alternative ways to improve performance when the property is unavailable. | Explain performance‑impact of XML map refresh in Aspose.Cells and list best‑practice techniques for large data writes. | Generate a robust try‑catch pattern for loading a workbook, performing high‑volume cell insertion, and handling missing EnableXMLMapRefresh functionality.

using System;
using System.IO;
using Aspose.Cells;

// This C# example loads an existing workbook, verifies the source file, writes values to 10,000 rows × 10 columns, and saves the result. It highlights that the EnableXMLMapRefresh property is not present in the current Aspose.Cells release, so the updates run with the default refresh behavior.
class DisableXmlMapRefreshDemo
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook wb = new Workbook(inputPath);

            // NOTE: The EnableXMLMapRefresh property is not available in this version of Aspose.Cells.
            // Bulk updates are performed without explicitly disabling XML map refresh.

            // Perform bulk cell updates
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue(row * col);
                }
            }

            // Save the modified workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
