// Title: Create a macro‑free Excel template workbook with predefined theme colors using Aspose.Cells for .NET
// AI Prompts: Generate a new macro‑free .xlsx workbook with Aspose.Cells for .NET and save it to a specified file path. | Insert code that sets custom theme colors in a freshly created workbook (when the theme API is available) before saving. | Wrap workbook creation and saving in a try‑catch block to capture exceptions and log error details in C#.
// Common Searches: how to generate a blank macro‑free Excel file with Aspose.Cells C# | Aspose.Cells create workbook and apply custom theme colors | save a template workbook as .xlsx using Aspose.Cells for .NET | which Aspose.Cells version supports theme customization | C# example for creating an Excel template without macros
// Tags: macro‑free workbook creation Aspose.Cells | Aspose.Cells predefined theme colors | save workbook as Xlsx C# | Aspose.Cells theme API version requirement | Excel template generation without macros

using System;
using System.Drawing;
using Aspose.Cells;

namespace TemplateGenerator
{
    // Demonstrates how to create a new macro‑free Workbook with Aspose.Cells for .NET, notes that applying predefined theme colors requires a newer version of the library, and saves the file as Template.xlsx with basic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new macro‑free workbook
                Workbook workbook = new Workbook();

                // NOTE: Theme manipulation APIs are not available in the referenced Aspose.Cells version.
                // If needed, upgrade Aspose.Cells to a version that supports Theme customization.
                // The following placeholder demonstrates where theme code would be inserted.

                // Save the template workbook to a file
                string outputPath = "Template.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
