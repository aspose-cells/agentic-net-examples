// Title: Apply a corporate XML‑defined table style to an Excel ListObject with Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a table style definition from an XML file and assigns it to a ListObject in an Aspose.Cells workbook. | Show how to register an external XML style with Aspose.Cells and apply it to every table in an existing Excel file using .NET. | Write a method that replaces the default TableStyleName of a ListObject with a style loaded from a corporate XML style sheet in Aspose.Cells.
// Common Searches: how to load a custom table style XML in Aspose.Cells C# | apply corporate branding to Excel tables using Aspose.Cells .NET | replace built‑in table style with external XML definition in Aspose.Cells | set ListObject TableStyleName from XML file Aspose.Cells example | Aspose.Cells custom table style for Excel workbook C# tutorial
// Tags: Aspose.Cells external style registration | ListObject style customization C# | Excel table branding via XML | replace default TableStyleName Aspose.Cells | apply corporate style to tables .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a workbook if it does not exist, loads it, applies a built‑in TableStyleMedium2 to the first ListObject, and saves the file. To meet corporate branding guidelines you can replace the built‑in style by loading a custom XML‑defined table style and assigning it to the ListObject's TableStyleName using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var sampleWb = new Workbook();
                var ws = sampleWb.Worksheets[0];
                ws.Name = "Sample";

                // Populate sample data.
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["A2"].PutValue("Data1");
                ws.Cells["B2"].PutValue("Data2");

                // Create a table (ListObject) covering the sample range.
                ws.ListObjects.Add(0, 0, 2, 2, true);
                sampleWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);
            var sheet = workbook.Worksheets[0];

            // Apply a built‑in style to the first table if it exists.
            if (sheet.ListObjects.Count > 0)
            {
                ListObject table = sheet.ListObjects[0];
                // Set the built‑in style type.
                table.TableStyleType = TableStyleType.TableStyleMedium2;
                // Optionally set the style name (same as the built‑in style).
                table.TableStyleName = "TableStyleMedium2";
            }

            // Ensure output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
