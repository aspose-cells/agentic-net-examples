// Title: Extract VBA Project Author from a Macro‑Enabled Workbook with Aspose.Cells for .NET
// Description: C# example that loads an .xlsm file using Aspose.Cells, verifies the presence of a VBA project with HasMacro, reads the "Author" built‑in document property, and prints the value while handling missing data.
// Keywords: Aspose.Cells VBA author | read .xlsm built‑in properties | Workbook.HasMacro C# | extract VBA metadata .NET | macro‑enabled workbook author | C# Aspose.Cells document properties
// Common Searches: how to get VBA author from xlsm using Aspose.Cells | read built‑in document properties of macro workbook C# | check if workbook contains VBA project Aspose.Cells | extract VBA project metadata .NET | Aspose.Cells get author of macro‑enabled Excel file
// Developer Intent: Obtain the author information stored in the VBA project of a macro‑enabled Excel workbook.
// Use Cases: Document ownership of macros for compliance audits. | Create a report of VBA authors across a collection of .xlsm files. | Validate the presence of a VBA project before applying transformations.
// AI Prompts: Generate C# code with Aspose.Cells that lists the author and other built‑in properties of a VBA project in an .xlsm file. | Write a method to scan a folder of .xlsm files, check HasMacro, and log each workbook's VBA author. | Build a utility that extracts VBA project metadata (author, title, comments) from a macro‑enabled workbook and exports it to JSON.

using System;
using Aspose.Cells;

// C# example that loads an .xlsm file using Aspose.Cells, verifies the presence of a VBA project with HasMacro, reads the "Author" built‑in document property, and prints the value while handling missing data.
class ExtractVbaAuthor
{
    static void Main()
    {
        // Load a macro-enabled workbook from file
        string inputPath = "input.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Verify that the workbook actually contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Retrieve the author metadata from the built‑in document properties
            var authorProperty = workbook.BuiltInDocumentProperties["Author"];
            string author = authorProperty != null && authorProperty.Value != null
                ? authorProperty.Value.ToString()
                : "Unknown";

            Console.WriteLine("VBA Project Author: " + author);
        }
        else
        {
            Console.WriteLine("The loaded workbook does not contain a VBA project.");
        }
    }
}
