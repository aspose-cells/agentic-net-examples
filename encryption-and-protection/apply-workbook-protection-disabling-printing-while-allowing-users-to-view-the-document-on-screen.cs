// Title: Disable printing while still allowing on‑screen viewing by protecting an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect workbook to prevent printing but allow view | how to block print option in an Excel file using Aspose.Cells .NET API | Aspose.Cells example that disables printing in a protected workbook | prevent users from printing an Excel workbook while still letting them view it with Aspose.Cells | set print restriction on an Excel workbook using Aspose.Cells for C#
// Tags: Aspose.Cells workbook protection disable printing | C# Aspose.Cells set print restriction | Excel file protect prevent printing .NET | Aspose.Cells protect workbook with password | disable Excel printing using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, applies full protection with a password using Aspose.Cells for .NET, which disables printing while still allowing the document to be viewed on screen, and then saves the protected file with error handling for missing input and runtime exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Protect the workbook with a password (all protection types)
            workbook.Protect(ProtectionType.All, "password");

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
