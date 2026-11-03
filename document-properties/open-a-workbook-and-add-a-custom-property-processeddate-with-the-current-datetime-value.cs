// Title: How to add a custom document property "ProcessedDate" with the current DateTime to an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add custom document property with current timestamp | set processed date property in Excel file using Aspose.Cells .NET | create or update custom property ProcessedDate in workbook programmatically | load existing workbook and add DateTime custom property with Aspose.Cells | save workbook after inserting custom document property in C#
// Tags: custom document property insertion Aspose.Cells | DateTime workbook metadata C# | Aspose.Cells workbook property manipulation | Excel file custom metadata .NET | programmatic workbook property update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one), adds a custom document property named "ProcessedDate" containing the current DateTime, and saves the file, demonstrating how to manage custom metadata with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                // If the workbook is password‑protected, you can specify the password via LoadOptions
                // LoadOptions loadOptions = new LoadOptions { Password = "yourPassword" };
                // workbook = new Workbook(inputPath, loadOptions);
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
            }

            // Add a custom document property named "ProcessedDate" with the current date and time
            workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

            // Save the workbook (overwrites the original file or creates a new one)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
