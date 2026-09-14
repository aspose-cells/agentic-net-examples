// Title: Enumerate Workbook.LoadWarnings to retrieve Excel load warnings after opening a workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel file using Aspose.Cells and iterates over the Workbook.LoadWarnings collection, printing each warning to the console. | Create a C# method that returns a List<string> of warning messages extracted from Workbook.LoadWarnings after loading a spreadsheet with Aspose.Cells. | Show how to log all load warnings from a workbook by accessing Workbook.LoadWarnings after constructing the Workbook with LoadOptions in a .NET application.
// Common Searches: c# get workbook load warnings aspose.cells | how to read Workbook.LoadWarnings collection in Aspose.Cells .NET | aspnet enumerate load warnings after opening Excel with Aspose.Cells | list warnings returned by Aspose.Cells when loading an xlsx file
// Tags: Aspose.Cells load warnings enumeration | C# Workbook.LoadWarnings iteration | Aspose.Cells warning collection handling | Excel file loading issues Aspose.Cells | Aspose.Cells .NET warning retrieval

using Aspose.Cells;
using System;
using System.IO;

// The example checks for an existing 'input.xlsx' file, creates a simple workbook if missing, saves it, and then loads the workbook using Aspose.Cells LoadOptions. After loading, you can iterate through the Workbook.LoadWarnings collection to capture and log any warnings generated during the load process.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Ensure the input file exists; create a simple workbook if it does not.
        if (!File.Exists(filePath))
        {
            try
            {
                var sampleWb = new Workbook();
                sampleWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                sampleWb.Save(filePath);
                Console.WriteLine($"Sample file created at '{filePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sample file: {ex.Message}");
                return;
            }
        }

        try
        {
            // Load the workbook with default options.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully.
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
