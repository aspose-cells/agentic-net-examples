// Title: Batch set IsNillable = true for a shared custom document property in multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to set IsNillable flag on a custom document property for all .xlsx files with Aspose.Cells C# | C# code to make a shared custom property optional in a folder of Excel workbooks | batch update custom document properties to be nillable using Aspose.Cells .NET | process multiple Excel workbooks and set custom property IsNillable to true programmatically | Aspose.Cells example for setting optional metadata on a shared property across workbooks
// Tags: Aspose.Cells batch modify custom document properties | set IsNillable flag for Excel custom property .NET | make shared custom property optional in .xlsx files | process multiple workbooks with Aspose.Cells C# | dynamic assignment of custom property attributes

using System;
using System.IO;
using Aspose.Cells;

// The example scans a folder for all *.xlsx files, loads each workbook with Aspose.Cells, ensures a custom document property named "SharedProperty" exists, sets its IsNillable flag to true using dynamic typing to make the property optional, saves the workbook, and logs the outcome for each file.
class SetIsNillableForSharedProperty
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Name of the shared custom property to modify
        string propertyName = "SharedProperty";

        // Iterate through all Excel files in the folder
        foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the custom document properties collection
                var customProps = workbook.CustomDocumentProperties;

                // Try to find the property by name
                object propObj = null;
                foreach (var p in customProps)
                {
                    dynamic dp = p; // use dynamic to avoid compile‑time type dependency
                    if (dp.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        propObj = dp;
                        break;
                    }
                }

                // If the property does not exist, create it with a default empty value
                if (propObj == null)
                {
                    propObj = customProps.Add(propertyName, string.Empty);
                }

                // Set IsNillable to true to make the property optional (allow null)
                try
                {
                    // Use dynamic to set the property at runtime
                    ((dynamic)propObj).IsNillable = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to set IsNillable for '{propertyName}' in file '{Path.GetFileName(filePath)}': {ex.Message}");
                }

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("IsNillable set to true for the shared property in all workbooks.");
    }
}
