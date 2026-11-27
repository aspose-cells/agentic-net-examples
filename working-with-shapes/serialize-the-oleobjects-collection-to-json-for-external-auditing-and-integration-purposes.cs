// Title: Export OleObject collection from an Excel workbook to JSON using Aspose.Cells in C#
// Common Searches: how to extract ole objects from an Excel file to json with Aspose.Cells C# | c# Aspose.Cells serialize embedded OLE object metadata to a JSON file | export worksheet oleobject properties to JSON using Aspose.Cells library | save oleobject details from a workbook as indented JSON in .NET | command line tool to dump Excel OLE objects to JSON with Aspose.Cells
// Tags: Aspose.Cells export OleObject metadata to JSON | C# serialize Excel OLE objects with Aspose.Cells | OleObject collection JSON dump Aspose.Cells | Excel embedded object extraction Aspose.Cells C# | Workbook OleObject properties JSON serialization

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectsJsonExport
{
    // The example loads an Excel workbook via Aspose.Cells, iterates each worksheet's OleObject collection, captures properties such as name, ProgID, source paths, positioning, and display settings into a custom POCO, then serializes the list to an indented JSON file. It supports command‑line arguments for input and output paths, creates missing directories, and includes basic error handling.
    class Program
    {
        public class OleObjectInfo
        {
            public string WorksheetName { get; set; }
            public int Index { get; set; }
            public string Name { get; set; }
            public string ProgID { get; set; }
            public string ObjectSourceFullName { get; set; }
            public string ImageSourceFullName { get; set; }
            public bool DisplayAsIcon { get; set; }
            public bool AutoLoad { get; set; }
            public bool AutoUpdate { get; set; }
            public string Label { get; set; }
            public string Title { get; set; }
            public string AlternativeText { get; set; }
            public int UpperLeftRow { get; set; }
            public int UpperLeftColumn { get; set; }
            public int LowerRightRow { get; set; }
            public int LowerRightColumn { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }

        static void Main(string[] args)
        {
            // Determine input and output paths.
            string inputPath;
            string outputPath;

            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }
            else
            {
                // Fallback to sample files in the current directory.
                inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.xlsx");
                outputPath = Path.Combine(Directory.GetCurrentDirectory(), "oleobjects.json");

                Console.WriteLine("Usage: OleObjectsJsonExport <inputWorkbookPath> <outputJsonPath>");
                Console.WriteLine($"Insufficient arguments supplied. Falling back to defaults:");
                Console.WriteLine($"  Input  : {inputPath}");
                Console.WriteLine($"  Output : {outputPath}");
            }

            // Validate input file existence.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input workbook not found at '{inputPath}'.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Prepare a list to hold serialized information.
                List<OleObjectInfo> oleInfos = new List<OleObjectInfo>();

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    OleObjectCollection oleCollection = sheet.OleObjects;

                    // Iterate through each OleObject in the collection.
                    for (int i = 0; i < oleCollection.Count; i++)
                    {
                        OleObject ole = oleCollection[i];

                        OleObjectInfo info = new OleObjectInfo
                        {
                            WorksheetName = sheet.Name,
                            Index = i,
                            Name = ole.Name,
                            ProgID = ole.ProgID,
                            ObjectSourceFullName = ole.ObjectSourceFullName,
                            ImageSourceFullName = ole.ImageSourceFullName,
                            DisplayAsIcon = ole.DisplayAsIcon,
                            AutoLoad = ole.AutoLoad,
                            AutoUpdate = ole.AutoUpdate,
                            Label = ole.Label,
                            Title = ole.Title,
                            AlternativeText = ole.AlternativeText,
                            UpperLeftRow = ole.UpperLeftRow,
                            UpperLeftColumn = ole.UpperLeftColumn,
                            LowerRightRow = ole.LowerRightRow,
                            LowerRightColumn = ole.LowerRightColumn,
                            Width = ole.Width,
                            Height = ole.Height
                        };

                        oleInfos.Add(info);
                    }
                }

                // Serialize the list to JSON with indentation for readability.
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(oleInfos, jsonOptions);

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to file.
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"Successfully exported {oleInfos.Count} OleObject(s) to JSON:");
                Console.WriteLine(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An unexpected error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
