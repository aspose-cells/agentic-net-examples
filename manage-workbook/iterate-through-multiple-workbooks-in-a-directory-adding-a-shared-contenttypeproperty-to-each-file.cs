// Title: Batch add or update a shared ContentTypeProperty in multiple Excel workbooks with Aspose.Cells for .NET (C#)
// Common Searches: C# code to add the same ContentTypeProperty to every .xlsx file in a folder using Aspose.Cells | How to loop through a directory of Excel workbooks and set a custom ContentTypeProperty with Aspose.Cells | Aspose.Cells batch update of custom properties across multiple Excel files in C# | Programmatically add or modify a shared ContentTypeProperty in many workbooks with Aspose.Cells .NET | Iterate over Excel files in a directory and overwrite them after adding a ContentTypeProperty using C#
// Tags: Aspose.Cells batch add ContentTypeProperty | C# update Excel custom properties multiple workbooks | process directory of .xlsx files Aspose.Cells | overwrite workbook after property modification Aspose.Cells | iterate Excel files folder C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies a target directory, enumerates .xlsx, .xls, and .xlsm files, loads each workbook with Aspose.Cells, adds or updates a ContentTypeProperty named "SharedProperty" with the value "SharedValue", saves the workbook back to the original file, and logs each processed file or any errors encountered.
class Program
{
    static void Main()
    {
        // Path to the directory containing the workbooks
        string folderPath = @"C:\Path\To\Workbooks";

        // Verify that the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Directory not found: {folderPath}");
            return;
        }

        // Define the shared ContentTypeProperty name and value
        string propertyName = "SharedProperty";
        string propertyValue = "SharedValue";

        // Retrieve all files in the directory
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string file in files)
        {
            // Process only Excel workbook files
            string ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                continue;

            // Ensure the file still exists before loading
            if (!File.Exists(file))
            {
                Console.WriteLine($"File not found (skipped): {file}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(file);

                // Add the shared ContentTypeProperty (or update if it already exists)
                var existingProp = workbook.ContentTypeProperties[propertyName];
                if (existingProp != null)
                {
                    existingProp.Value = propertyValue;
                }
                else
                {
                    workbook.ContentTypeProperties.Add(propertyName, propertyValue);
                }

                // Save the workbook, overwriting the original file
                workbook.Save(file);
                Console.WriteLine($"Processed: {file}");
            }
            catch (Exception ex)
            {
                // Log any errors for this file and continue with the next one
                Console.WriteLine($"Error processing file '{file}': {ex.Message}");
            }
        }
    }
}
