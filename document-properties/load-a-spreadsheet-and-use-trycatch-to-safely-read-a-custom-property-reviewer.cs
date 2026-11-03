// Title: How to load an Excel workbook and safely read the custom property "Reviewer" with try‑catch in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells read custom document property with error handling | How to retrieve the Reviewer custom property from an Excel file using Aspose.Cells | Try-catch example for loading a workbook and accessing custom properties in Aspose.Cells | Check for existence of a custom property before reading it with Aspose.Cells C# | Handle missing custom property exception when using Aspose.Cells workbook
// Tags: read custom document property Aspose.Cells | load workbook with try-catch C# | handle missing custom property Aspose | Excel custom property extraction C# | exception handling Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook inside a try‑catch block, and then attempts to read the custom document property named "Reviewer" within another try‑catch, outputting the value or appropriate error messages if the property is missing or an exception occurs.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Retrieve the custom property named "Reviewer"
            var reviewerProp = workbook.CustomDocumentProperties["Reviewer"];
            if (reviewerProp != null && reviewerProp.Value != null)
            {
                Console.WriteLine("Reviewer: " + reviewerProp.Value);
            }
            else
            {
                Console.WriteLine("Custom property \"Reviewer\" not found.");
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur while accessing the property
            Console.WriteLine("Error reading custom property \"Reviewer\": " + ex.Message);
        }
    }
}
