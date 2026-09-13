// Title: Set a shared custom document property to null (make it nillable) in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code with Aspose.Cells that opens a workbook, finds a shared custom document property, assigns it a null value to make it optional, and saves the file. | Write a script that creates or loads an Excel workbook, checks for a specific custom property across the workbook, sets its value to null (nillable), and persists the changes with Aspose.Cells.
// Common Searches: Aspose.Cells C# make custom document property optional by setting it to null | how to clear a shared custom property value in an Excel workbook using Aspose.Cells .NET | set shared custom document property to null across multiple workbooks with Aspose.Cells
// Tags: Aspose.Cells set custom document property null | C# make shared workbook metadata optional | clear custom document property value Aspose.Cells | modify Excel custom properties programmatically | optional custom document property .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a Workbook, accesses the custom document property named "SharedProperty", sets its Value to null (making the property nillable) if it exists, and saves the workbook as ModifiedWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                var workbook = new Workbook();

                // Access the custom document property named "SharedProperty".
                // Replace "SharedProperty" with the actual property name you want to modify.
                var sharedProperty = workbook.CustomDocumentProperties["SharedProperty"];

                // If the property exists, set its value to null (making it effectively nillable).
                if (sharedProperty != null)
                {
                    sharedProperty.Value = null;
                }

                // Define the output file path.
                string outputPath = "ModifiedWorkbook.xlsx";

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
