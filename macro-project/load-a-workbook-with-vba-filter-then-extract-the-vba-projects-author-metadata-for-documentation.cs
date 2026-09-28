// Title: Extract author metadata from a macro‑enabled Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsm file with Aspose.Cells, checks whether a VBA project exists, and returns the workbook's built‑in Author property while safely handling missing files. | Create a reusable C# method that loads a macro‑enabled workbook, validates the presence of VBA macros, extracts the Author document property, and logs appropriate error messages.
// Common Searches: Aspose.Cells C# get Author property from macro enabled Excel file | How to verify VBA macros exist in an .xlsm workbook using Aspose.Cells | Read built‑in document properties of .xlsm with Aspose.Cells .NET | C# code to handle FileNotFoundException when opening Excel workbook with Aspose.Cells | Extract workbook author when VBA project has no author field Aspose.Cells
// Tags: retrieve built-in Author property from .xlsm using Aspose.Cells | detect VBA project presence in Excel workbook with Aspose.Cells | load macro-enabled workbook and read document properties .NET | handle file-not-found errors when opening Excel with Aspose.Cells | use VbaProject object to verify macro existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads a macro‑enabled .xlsm workbook with Aspose.Cells, checks for the presence of a VBA project, and then reads the workbook's built‑in Author document property (since the VbaProject object lacks an Author field). It outputs the author value or a default placeholder and includes robust handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook that may contain VBA macros
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // The VbaProject class does not expose an Author property.
                // Retrieve the workbook's built‑in Author property instead.
                string workbookAuthor = workbook.BuiltInDocumentProperties["Author"]?.ToString() ?? "Unknown";
                Console.WriteLine("Workbook Author: " + workbookAuthor);
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
