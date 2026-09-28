// Title: Batch add a shared custom document property to all .xlsx workbooks in a folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates every .xlsx file in a given directory, loads each workbook with Aspose.Cells, checks for a custom document property named 'SharedProperty', adds it with value 'SharedValue' when missing, and saves the file. | Create a robust .NET routine that processes a folder of Excel files, injects a shared custom document property via Aspose.Cells, includes existence check and exception handling, and overwrites the original workbooks.
// Common Searches: Aspose.Cells C# add same custom document property to multiple Excel files in a folder | how to batch update custom document properties in .xlsx workbooks using .NET | iterate through directory of workbooks and set shared property with Aspose.Cells API
// Tags: batch add custom document property Aspose.Cells | enumerate .xlsx files C# Aspose.Cells | shared custom property across workbooks .NET | custom document properties automation Aspose.Cells | process Excel workbooks in folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, adds a custom document property named 'SharedProperty' with value 'SharedValue' if it does not already exist, and saves the workbook back to its original location.
class Program
{
    static void Main()
    {
        try
        {
            // Directory containing the workbooks
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Name and value of the shared custom document property to add
            const string propertyName = "SharedProperty";
            const string propertyValue = "SharedValue";

            // Get all Excel files in the directory
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Add the shared custom document property if it does not already exist
                    if (!workbook.CustomDocumentProperties.Contains(propertyName))
                    {
                        // Add a new custom document property directly
                        workbook.CustomDocumentProperties.Add(propertyName, propertyValue);
                    }

                    // Save the workbook (overwrites the original file)
                    workbook.Save(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Custom document property added to all workbooks.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
