// Title: C# command‑line utility to add or update a custom document property in an Excel workbook with Aspose.Cells and ensure proper disposal
// Common Searches: how to add or update a custom document property in an Excel file using Aspose.Cells C# | C# console application to modify Excel custom properties with Aspose.Cells | saving workbook after changing a custom document property with Aspose.Cells | disposing Aspose.Cells workbook in a command‑line tool | command‑line editor for Excel custom properties using Aspose.Cells
// Tags: add custom document property Aspose.Cells | update Excel custom property C# | save workbook after property modification Aspose.Cells | dispose Aspose.Cells workbook resources | command‑line Excel property editor Aspose.Cells

using System;
using Aspose.Cells;

// A C# console program that loads an Excel file, optionally adds or updates a custom document property via Aspose.Cells, saves the workbook to persist the change, and then disposes the workbook to release all resources.
class Program
{
    static void Main(string[] args)
    {
        // Verify that at least the Excel file path is provided
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: tool <excelFilePath> [propertyName propertyValue]");
            return;
        }

        string filePath = args[0];

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // If a property name and value are supplied, add or update the custom document property
        if (args.Length >= 3)
        {
            string propName = args[1];
            string propValue = args[2];

            var customProps = workbook.CustomDocumentProperties;

            // Update existing property if it exists, otherwise add a new one
            var existingProp = customProps[propName];
            if (existingProp != null)
            {
                existingProp.Value = propValue;
            }
            else
            {
                customProps.Add(propName, propValue);
            }

            // Save the workbook to persist the new property
            workbook.Save(filePath);
        }

        // Release all resources used by the workbook
        workbook.Dispose();
    }
}
