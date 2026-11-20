// Title: Disable the LinksUpToDate built‑in document property in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set LinksUpToDate property to false in C# | prevent Excel workbook from checking external links with Aspose.Cells | disable automatic link validation when saving .xlsx using Aspose.Cells .NET | how to turn off LinksUpToDate built‑in property in C# Aspose.Cells example | modify built‑in document properties to stop link updates in Excel via Aspose.Cells | C# code to suppress external link checks in an Excel file using Aspose.Cells
// Tags: Aspose.Cells LinksUpToDate property | disable automatic link checks Excel | C# built‑in document property modification | Excel workbook link update suppression | Aspose.Cells workbook save without link validation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook (or creates a new one if missing), accesses the built‑in document property named "LinksUpToDate", sets its value to false to stop automatic external‑link validation, and then saves the workbook, handling any exceptions that may occur during loading, property modification, or saving.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputFile))
                {
                    workbook = new Workbook(inputFile);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                try
                {
                    // Disable automatic link updates by setting the built‑in property.
                    var linksProp = workbook.BuiltInDocumentProperties["LinksUpToDate"];
                    if (linksProp != null)
                    {
                        linksProp.Value = false;
                    }
                }
                catch (Exception propEx)
                {
                    Console.WriteLine($"Failed to set LinksUpToDate property: {propEx.Message}");
                }

                // Save the workbook.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
