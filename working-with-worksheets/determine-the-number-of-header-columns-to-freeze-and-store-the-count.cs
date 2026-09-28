// Title: How to get the count of frozen header columns in an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells and returns the number of frozen columns in the first worksheet, handling both View.FreezePanesColumn and ViewOptions.FreezePanesColumn. | Create a reusable method that accepts a Worksheet object and returns the frozen column count, including a fallback for older Aspose.Cells versions where the View property is unavailable. | Extend the example to also display the frozen row count together with the frozen column count using Aspose.Cells.
// Common Searches: aspnet get frozen column count from Excel using Aspose.Cells | c# read freeze panes column property Aspose.Cells version compatibility | how to determine number of frozen header columns in a worksheet with Aspose.Cells | retrieve freeze panes settings rows and columns from Excel file using Aspose.Cells .NET | fallback to ViewOptions when View property missing in Aspose.Cells
// Tags: Aspose.Cells FreezePanesColumn retrieval | C# fallback ViewOptions FreezePanesColumn | Excel worksheet frozen columns count | Aspose.Cells version compatibility freeze panes | read frozen rows and columns Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

namespace AsposeCellsExample
{
    // // Loads an Excel workbook, accesses the first worksheet, and reads the number of frozen header columns via View.FreezePanesColumn with a fallback to ViewOptions.FreezePanesColumn for older Aspose.Cells versions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Use dynamic to access view properties (covers different Aspose.Cells versions)
                dynamic wsDynamic = worksheet;
                int frozenHeaderColumns = 0;

                try
                {
                    frozenHeaderColumns = (int)wsDynamic.View.FreezePanesColumn;
                }
                catch (RuntimeBinderException)
                {
                    // Fallback for versions where the View property is unavailable
                    try
                    {
                        frozenHeaderColumns = (int)wsDynamic.ViewOptions.FreezePanesColumn;
                    }
                    catch
                    {
                        // If neither property exists, default to 0
                        frozenHeaderColumns = 0;
                    }
                }

                Console.WriteLine($"Number of frozen header columns: {frozenHeaderColumns}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
