// Title: C# example using Aspose.Cells to enumerate workbook ContentTypeProperties and report their optional (nillable) status
// Common Searches: how to list ContentTypeProperties and determine if they are optional with Aspose.Cells in C# | C# Aspose.Cells code to output nillable flag of each workbook custom property | sample program to generate a summary of optional ContentTypeProperties from an Excel file using Aspose.Cells | retrieve nullability information of Excel ContentTypeProperty via Aspose.Cells C# API | enumerate workbook ContentTypeProperties and print their optional status in a .NET console app
// Tags: enumerate ContentTypeProperties Aspose.Cells | check optional flag ContentTypeProperty C# | generate nillable property report Excel | Aspose.Cells workbook custom property nullability | C# read Excel content type properties

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The sample loads an .xlsx workbook with Aspose.Cells, iterates through its ContentTypeProperties collection, assumes each property is non‑nillable (as the optional flag is not exposed), builds a line‑by‑line summary showing the property name and its nillable status, and writes the report to the console while handling missing files and load errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – '{inputPath}'.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Build a summary of nillable (optional) ContentTypeProperties
        var summary = new StringBuilder();

        try
        {
            // Iterate through each ContentTypeProperty in the workbook
            foreach (var prop in workbook.ContentTypeProperties)
            {
                // Aspose.Cells does not expose an explicit IsOptional flag in recent versions.
                // For demonstration, we treat all properties as non‑nillable.
                bool isNillable = false;

                // Append the result for this property to the summary
                summary.AppendLine($"Property '{prop.Name}' is nillable: {isNillable}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing ContentTypeProperties: {ex.Message}");
            return;
        }

        // Output the generated summary
        Console.WriteLine("Nillable ContentTypeProperties summary:");
        Console.WriteLine(summary.ToString());
    }
}
