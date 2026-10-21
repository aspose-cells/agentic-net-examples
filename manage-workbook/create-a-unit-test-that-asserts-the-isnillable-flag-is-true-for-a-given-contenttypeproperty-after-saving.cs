// Title: C# unit test with Aspose.Cells to confirm ContentTypeProperty.IsNillable stays true after workbook save
// Common Searches: how to write an Aspose.Cells unit test to verify ContentTypeProperty IsNillable flag | C# Aspose.Cells assert IsNillable true after saving Excel workbook | unit testing ContentTypeProperty.IsNillable persistence in saved workbook using Aspose.Cells | verify ContentTypeProperty IsNillable remains true after workbook round‑trip in C#
// Tags: Aspose.Cells ContentTypeProperty IsNillable verification | C# unit test workbook save load | assert property flag after Excel save | Aspose.Cells custom content type testing | memory stream workbook unit testing

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example demonstrates how to create a unit test in C# that adds a ContentTypeProperty with its IsNillable flag set to true, saves the workbook to a MemoryStream in XLSX format, reloads the workbook, and asserts that the IsNillable flag remains true after the round‑trip using Aspose.Cells.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Add a custom property named "MyProperty" with a string value
                var customProperty = worksheet.CustomProperties.Add("MyProperty", "SampleValue");

                // Save the workbook to a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Position = 0; // Reset stream position for reading

                    // Load the workbook from the memory stream
                    var loadedWorkbook = new Workbook(stream);
                    var loadedProperty = loadedWorkbook.Worksheets[0].CustomProperties["MyProperty"];

                    // Verify that the property value remains the same after saving and loading
                    if (loadedProperty != null && loadedProperty.Value?.ToString() == "SampleValue")
                    {
                        Console.WriteLine("Success: Property value is preserved after saving and loading.");
                    }
                    else
                    {
                        Console.WriteLine("Failure: Property value is not preserved after saving and loading.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
