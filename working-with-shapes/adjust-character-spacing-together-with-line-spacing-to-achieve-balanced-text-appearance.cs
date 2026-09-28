// Title: Increase line spacing by setting row height and emulate character spacing in Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to enlarge the vertical space between lines in a cell by setting the row height to a specific point value. | Explain a workaround to mimic text tracking in Aspose.Cells for .NET when the API lacks a direct property. | Show how to combine row‑height changes with additional font styling such as bold or italic using Aspose.Cells.
// Common Searches: Aspose.Cells C# increase line spacing by changing row height | how to simulate character spacing in Aspose.Cells .NET | set row height for better text appearance in Excel using Aspose.Cells | Aspose.Cells adjust cell style when character spacing property is missing | increase row height to affect line spacing in Aspose.Cells workbook
// Tags: row dimension adjustment Aspose.Cells .NET | spacing via row size Aspose.Cells | font tracking workaround Aspose.Cells | cell style modification Aspose.Cells C# | font formatting options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates or loads an Excel workbook, accesses cell A1, clones its style (noting that Aspose.Cells does not expose a character‑spacing property), sets the row height to 25 points to increase line spacing, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; create a blank workbook if it does not
                if (!File.Exists(inputPath))
                {
                    var tempWorkbook = new Workbook();
                    tempWorkbook.Save(inputPath);
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Choose the cell to format (e.g., A1)
                var cell = worksheet.Cells["A1"];

                // Clone the existing style
                var style = cell.GetStyle();

                // NOTE: Aspose.Cells does not provide a direct property for character spacing.
                // If needed, other font properties (e.g., Bold, Italic) can be set here.
                // Example: style.Font.IsBold = true;

                // Apply the (potentially modified) style back to the cell
                cell.SetStyle(style);

                // Adjust line spacing by modifying the row height.
                var row = worksheet.Cells.Rows[cell.Row];
                row.Height = 25; // Set row height to 25 points

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
