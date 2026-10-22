// Title: How to update a linked cell range and refresh a picture shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that changes the values of cells A1:B2 in a worksheet and then refreshes the picture shape named "Picture 1" with Aspose.Cells, including fallback handling if the Update method is unavailable. | Demonstrate how to invoke Picture.Update dynamically on a shape after modifying its linked cells, with error handling for older Aspose.Cells versions.
// Common Searches: Aspose.Cells C# refresh picture shape after editing linked cells | programmatically update a picture linked to a cell range in an Excel file using Aspose.Cells | dynamic call to Picture.Update when method is missing in Aspose.Cells .NET
// Tags: refresh picture shape Aspose.Cells C# | update linked cell range Aspose.Cells | dynamic Picture.Update invocation | handle missing Picture.Update Aspose.Cells | modify cells and refresh picture Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads an existing workbook, updates cells A1‑B2, retrieves the shape named "Picture 1", attempts to refresh the picture by calling Picture.Update via dynamic invocation (with fallback for older versions), and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Update the cell range that is linked to the picture
                sheet.Cells["A1"].PutValue("Updated Value 1");
                sheet.Cells["A2"].PutValue("Updated Value 2");
                sheet.Cells["B1"].PutValue("Updated Value 3");
                sheet.Cells["B2"].PutValue("Updated Value 4");

                // Retrieve the picture shape that is linked to the above range
                Shape shape = sheet.Shapes["Picture 1"];
                if (shape is Picture picture)
                {
                    // Attempt to refresh the picture; use dynamic to avoid compile‑time dependency on Update()
                    try
                    {
                        ((dynamic)picture).Update();
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // Update method not available in this version of Aspose.Cells; ignore
                    }
                }
                else
                {
                    Console.WriteLine("The specified shape is not a picture or does not exist.");
                }

                // Save the workbook with the refreshed picture (if possible)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
