// Title: C# unit test to confirm that a custom ContentTypeProperty’s IsNillable flag is retained after saving an Aspose.Cells workbook to XLSX
// Common Searches: how to assert IsNillable flag on ContentTypeProperty after saving workbook with Aspose.Cells C# | unit testing persistence of custom content type property in Aspose.Cells | verify that dynamic ContentTypeProperties retain IsNillable in saved Excel file | C# Aspose.Cells test for ContentTypeProperty IsNillable after workbook round‑trip | checking IsNillable flag of added content type property when reloading workbook
// Tags: Aspose.Cells ContentTypeProperty IsNillable persistence | C# unit test workbook save and reload | dynamic ContentTypeProperties in Excel workbook | custom content type property verification Aspose.Cells | assert property flag after workbook roundtrip

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a C# unit‑style test that adds a custom ContentTypeProperty to a Workbook, sets its IsNillable flag to true, saves the workbook to a MemoryStream in XLSX format, reloads it, and asserts that the property still exists and the IsNillable flag remains true.
    public class ContentTypePropertyTests
    {
        // Test method without NUnit dependencies
        public static void IsNillableFlag_ShouldBeTrue_AfterSavingWorkbook()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Use dynamic to work with ContentTypeProperties without compile‑time dependency on ContentTypeProperty type
                dynamic ctProps = workbook.ContentTypeProperties;

                // Add a custom content type property
                dynamic prop = ctProps.Add("TestProperty", "TestValue");

                // Set the IsNillable flag
                prop.IsNillable = true;

                // Save the workbook to a memory stream (no file I/O)
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0; // Reset stream position for reading

                    // Load the workbook back from the memory stream
                    var loadedWorkbook = new Workbook(ms);

                    // Retrieve the same property by name
                    dynamic loadedCtProps = loadedWorkbook.ContentTypeProperties;
                    dynamic loadedProp = loadedCtProps["TestProperty"];

                    // Validate that the property exists and the IsNillable flag persisted
                    if (loadedProp == null)
                        throw new InvalidOperationException("ContentTypeProperty should exist after loading.");

                    if (!loadedProp.IsNillable)
                        throw new InvalidOperationException("IsNillable flag should be true after saving and loading.");
                }

                Console.WriteLine("Test passed: IsNillable flag persisted correctly.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Run the test
            ContentTypePropertyTests.IsNillableFlag_ShouldBeTrue_AfterSavingWorkbook();
        }
    }
}
