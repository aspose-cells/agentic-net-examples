// Title: Batch update the custom "Version" document property to 3.0 in all Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells iterate over a folder and set custom document property Version to 3.0 | How to bulk modify a custom property in multiple .xls and .xlsx files with Aspose.Cells | Programmatically change the Version property of many Excel workbooks in a directory using .NET | Aspose.Cells example for updating custom document properties across several Excel files | C# code to set the same custom property value for all Excel files in a folder
// Tags: batch update custom document property Aspose.Cells | set version property multiple Excel workbooks C# | iterate Excel files folder Aspose.Cells | modify custom properties .xls .xlsx programmatically | Aspose.Cells bulk document property change | C# update Excel workbook custom property

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a specified directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, ensures a custom document property named "Version" exists, sets its value to "3.0", saves the workbook back in its original format, and logs the outcome for each file.
class UpdateDocumentVersion
{
    static void Main()
    {
        try
        {
            // Path to the directory containing Excel files
            string folderPath = @"C:\Path\To\ExcelFiles";

            // Verify that the directory exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files (both .xls and .xlsx) in the directory
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

            foreach (string filePath in excelFiles)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Set a custom document property named "Version"
                    // If the property already exists, update its value
                    const string versionPropertyName = "Version";
                    const string versionValue = "3.0";

                    var customProps = workbook.CustomDocumentProperties;
                    if (customProps.Contains(versionPropertyName))
                    {
                        // Update existing property value
                        customProps[versionPropertyName].Value = versionValue;
                    }
                    else
                    {
                        // Add new custom property
                        customProps.Add(versionPropertyName, versionValue);
                    }

                    // Save the workbook back to the same file, preserving the original format
                    workbook.Save(filePath);
                    Console.WriteLine($"Updated version for: {Path.GetFileName(filePath)}");
                }
                catch (Exception exFile)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
