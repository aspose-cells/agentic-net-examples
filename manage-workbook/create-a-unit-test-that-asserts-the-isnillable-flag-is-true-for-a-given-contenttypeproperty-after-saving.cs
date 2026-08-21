// Title: C# Unit Test: Verify ContentTypeProperty.IsNillable Persists After Workbook Save/Load
// Description: Creates a Workbook, adds a custom ContentTypeProperty with IsNillable set to true, saves to a temporary .xlsx file, reloads the file, and asserts that the IsNillable flag remains true. Includes cleanup of the temporary file.
// Keywords: Aspose.Cells | ContentTypeProperty | IsNillable | unit test | C# | .NET | Excel workbook serialization | MSTest | NUnit | xUnit
// Common Searches: Aspose.Cells unit test IsNillable | assert ContentTypeProperty after saving workbook | C# test property persistence in Excel file | how to verify custom content type property in Aspose.Cells
// Developer Intent: Write an automated test that confirms the IsNillable flag of a ContentTypeProperty stays true after the workbook is saved and reloaded.
// Use Cases: Ensure custom metadata flags survive workbook serialization for compliance reporting. | Add regression coverage for Aspose.Cells updates that might affect property handling. | Integrate into CI pipelines to validate Excel metadata integrity in enterprise solutions.
// AI Prompts: Generate an MSTest method that adds a ContentTypeProperty with IsNillable = true, saves the workbook, reloads it, and asserts the flag is true. | Create an NUnit test for Aspose.Cells that verifies IsNillable persists after workbook serialization. | Provide an xUnit test example that checks the loaded ContentTypeProperty.IsNillable matches the value set before saving.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsTests
{
    // Creates a Workbook, adds a custom ContentTypeProperty with IsNillable set to true, saves to a temporary .xlsx file, reloads the file, and asserts that the IsNillable flag remains true. Includes cleanup of the temporary file.
    class Program
    {
        static void Main()
        {
            string tempFile = null;

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a content type property and set IsNillable to true
                workbook.ContentTypeProperties.Add("Admin", "Aspose", "text");
                ContentTypeProperty property = workbook.ContentTypeProperties["Admin"];
                property.IsNillable = true;

                // Save to a temporary file
                tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempFile);

                // Ensure the file was created
                if (!File.Exists(tempFile))
                {
                    Console.WriteLine("Failed to create the temporary workbook file.");
                    return;
                }

                // Load the workbook from the saved file
                Workbook loadedWorkbook = new Workbook(tempFile);

                // Retrieve the property after loading
                ContentTypeProperty loadedProperty = loadedWorkbook.ContentTypeProperties["Admin"];

                // Verify that IsNillable is still true
                if (loadedProperty.IsNillable)
                {
                    Console.WriteLine("Success: IsNillable flag is true after saving and loading.");
                }
                else
                {
                    Console.WriteLine("Failure: IsNillable flag is false after saving and loading.");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary file
                if (!string.IsNullOrEmpty(tempFile) && File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // Suppress any cleanup exceptions
                    }
                }
            }
        }
    }
}
