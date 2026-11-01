// Title: C# console application using Aspose.Cells to load an Excel workbook, optionally add a custom document property, and ensure proper disposal
// Common Searches: c# console program to load an Excel file with Aspose.Cells and add a custom document property | how to add or update a custom document property in an existing workbook using Aspose.Cells CLI | ensure Aspose.Cells workbook is disposed correctly in a C# command line application | aspose.cells command line tool for adding custom properties to Excel spreadsheets | example of using Aspose.Cells in a C# console app to modify workbook metadata
// Tags: custom document property addition Aspose.Cells | load workbook from file Aspose.Cells | workbook disposal pattern Aspose.Cells | C# console Excel modification utility | update workbook metadata Aspose.Cells | command line Excel processing Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCommandLine
{
    // A C# console utility that accepts a file path, loads the Excel workbook with Aspose.Cells, optionally adds or updates a custom document property based on additional arguments, and guarantees the workbook is disposed in a finally block to release resources.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that at least the file path is provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsCommandLine <filePath> [propertyName propertyValue]");
                return;
            }

            string filePath = args[0];

            // Load the workbook from the specified file
            Workbook workbook = null;
            try
            {
                workbook = new Workbook(filePath);

                // If a property name and value are supplied, add them as a custom document property
                if (args.Length >= 3)
                {
                    string propertyName = args[1];
                    string propertyValue = args[2];

                    // Add or update the custom property
                    // If the property already exists, remove it first to avoid duplication
                    var existingProp = workbook.CustomDocumentProperties[propertyName];
                    if (existingProp != null)
                    {
                        workbook.CustomDocumentProperties.Remove(propertyName);
                    }

                    workbook.CustomDocumentProperties.Add(propertyName, propertyValue);
                    Console.WriteLine($"Added custom property: {propertyName} = {propertyValue}");
                }
                else
                {
                    Console.WriteLine("No custom property supplied; workbook loaded without modifications.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
            finally
            {
                // Ensure the workbook is properly disposed to release resources
                if (workbook != null)
                {
                    workbook.Dispose();
                }
            }
        }
    }
}
