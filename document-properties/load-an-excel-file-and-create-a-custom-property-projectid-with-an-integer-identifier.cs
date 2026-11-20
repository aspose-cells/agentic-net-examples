// Title: Create or update an integer custom document property named ProjectId in an Excel workbook with Aspose.Cells for .NET
// Common Searches: C# code to add an integer custom document property called ProjectId to an existing .xlsx using Aspose.Cells | How to set a custom property value in an Excel file with the Aspose.Cells .NET library | Aspose.Cells example for adding or updating custom document properties in a workbook | Saving custom properties to a workbook after loading or creating it with Aspose.Cells | Add ProjectId metadata to an Excel workbook programmatically in C#
// Tags: Workbook.CustomDocumentProperties.Add method | ProjectId integer property Aspose.Cells | Excel workbook metadata insertion .NET | Aspose.Cells custom properties API | programmatic addition of custom metadata to .xlsx

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel file (or creates a new workbook if the file is missing), adds a custom document property named "ProjectId" with the integer value 12345 using the Workbook.CustomDocumentProperties.Add method, and saves the workbook back to the same path.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file
            string inputPath = "input.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook
                }

                // Add a custom document property named "ProjectId" with an integer value
                // Use the overload that accepts name and value (isLinkToContent defaults to false)
                workbook.CustomDocumentProperties.Add("ProjectId", 12345);

                // Save the workbook (overwrites the original file or creates a new one)
                workbook.Save(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
