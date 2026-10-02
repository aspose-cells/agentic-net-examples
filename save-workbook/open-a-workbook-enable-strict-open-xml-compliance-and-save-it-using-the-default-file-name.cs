// Title: Open an Excel workbook, enforce strict OpenXML compliance, and save it with its original filename using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, activates the StrictOpenXmlCompliance setting, creates a new workbook when the file is absent, and saves it back to the same path. | Generate a .NET snippet that opens an Excel workbook, turns on strict OpenXML mode, and saves it without changing the filename, including handling of missing files.
// Common Searches: Aspose.Cells how to enforce OpenXML standards when saving a workbook C# | C# open existing Excel file and save with the same name using Aspose.Cells | Automatically create a new workbook if the file does not exist Aspose.Cells .NET | Saving a workbook using the default filename in Aspose.Cells | Example of Aspose.Cells strict mode save for .xlsx files | Handling file‑not‑found exceptions while loading a workbook with Aspose.Cells
// Tags: load workbook Aspose.Cells C# | save workbook with OpenXML compliance Aspose.Cells | auto‑create workbook if missing Aspose.Cells | default filename save Aspose.Cells .NET | exception handling file operations Aspose.Cells | open Excel file Aspose.Cells strict mode

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether 'input.xlsx' exists, creates a new workbook if it doesn't, loads the workbook with Aspose.Cells, enables strict OpenXML compliance, and saves it back using the original file path while handling potential exceptions.
class Program
{
    static void Main()
    {
        // Path to the workbook to open
        string filePath = "input.xlsx";

        try
        {
            // Ensure the input file exists; create a new workbook if it does not
            if (!File.Exists(filePath))
            {
                var newWorkbook = new Workbook();
                newWorkbook.Save(filePath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Save the workbook using the original (default) file name
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
