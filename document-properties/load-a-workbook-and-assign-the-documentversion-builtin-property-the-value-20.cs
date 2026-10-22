// Title: Set the DocumentVersion built‑in property of an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set built‑in DocumentVersion property in Excel file | programmatically update Excel workbook DocumentVersion using Aspose.Cells | how to assign version number to an .xlsx file with Aspose.Cells .NET | C# code to modify BuiltInDocumentProperties DocumentVersion in existing workbook
// Tags: Aspose.Cells set BuiltInDocumentProperties DocumentVersion | C# update Excel built‑in property with Aspose | modify workbook DocumentVersion property programmatically | Aspose.Cells document version metadata | Excel file versioning using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx or creates a new workbook, accesses the BuiltInDocumentProperties collection, sets the DocumentVersion property to "2.0", and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Set the built‑in property "DocumentVersion" to "2.0" if it exists.
            var docVersionProp = workbook.BuiltInDocumentProperties["DocumentVersion"];
            if (docVersionProp != null)
            {
                docVersionProp.Value = "2.0";
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
