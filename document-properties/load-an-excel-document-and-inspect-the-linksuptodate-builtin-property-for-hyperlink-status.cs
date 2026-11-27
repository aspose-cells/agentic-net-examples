// Title: How to read the LinksUpToDate built‑in property of an Excel workbook with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# read LinksUpToDate built‑in property from Excel file | Get hyperlink update status of an Excel workbook using Aspose.Cells .NET | Check if external links are up to date in an Excel file with Aspose.Cells API | Convert Aspose.Cells built‑in property value to Boolean in C# code | Load workbook and retrieve LinksUpToDate flag using Aspose.Cells library
// Tags: Aspose.Cells read built‑in document property | LinksUpToDate property extraction C# | Excel hyperlink status check Aspose.Cells | Workbook property conversion to boolean | C# load Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of input.xlsx, loads it with Aspose.Cells, accesses the built‑in document property "LinksUpToDate", safely converts its value (bool, int, or string) to a Boolean, and prints the hyperlink‑status flag to the console.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: File not found -> {inputFile}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the Excel workbook from the specified file
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Retrieve the built‑in property "LinksUpToDate"
        var linksProp = workbook.BuiltInDocumentProperties["LinksUpToDate"];

        // Determine the boolean status of the property
        bool linksUpToDate = false;
        if (linksProp != null && linksProp.Value != null)
        {
            // The property may be stored as bool, int (0/1) or string
            if (linksProp.Value is bool b)
                linksUpToDate = b;
            else if (linksProp.Value is int i)
                linksUpToDate = i != 0;
            else
            {
                try
                {
                    linksUpToDate = Convert.ToBoolean(linksProp.Value);
                }
                catch
                {
                    // If conversion fails, keep default false
                }
            }
        }

        // Output the status
        Console.WriteLine($"LinksUpToDate: {linksUpToDate}");
    }
}
