// Title: Load Excel template from byte array, bind smart marker data source, and retrieve processed workbook as byte array with Aspose.Cells for .NET
// AI Prompts: Load a workbook from a MemoryStream, attach an ArrayList of Person objects to the 'Persons' smart marker via WorkbookDesigner, invoke processing, and obtain the resulting file as a byte array. | Given a byte[] containing an Excel template, initialize WorkbookDesigner, set a custom data source, run Process, and capture the updated workbook in another byte array without writing to disk.
// Common Searches: aspnet load excel template from byte array for smart marker replacement | binding a collection to a smart marker using WorkbookDesigner in C# | convert processed Aspose.Cells workbook to byte array without saving to disk | how to use MemoryStream with Aspose.Cells to apply smart markers
// Tags: Aspose.Cells WorkbookDesigner smart marker binding | read Excel template into memory from byte array C# | execute smart marker processing in memory | export workbook to byte array Aspose.Cells | associate Person POCO list with smart marker named Persons

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace SmartMarkerApp
{
    // The example shows how to load an Excel template from a byte array using a MemoryStream, bind an ArrayList of Person POCO objects to a smart marker called 'Persons' with WorkbookDesigner, process all smart markers, and return the modified workbook as a byte array, enabling in‑memory template manipulation without file I/O.
    public class SmartMarkerProcessor
    {
        // Loads a workbook template from a byte array, applies smart marker data,
        // processes the markers, and returns the resulting workbook as a byte array.
        public static byte[] ProcessTemplate(byte[] templateBytes)
        {
            try
            {
                using (MemoryStream templateStream = new MemoryStream(templateBytes))
                {
                    Workbook workbook = new Workbook(templateStream);

                    // Set up the WorkbookDesigner to work with smart markers.
                    WorkbookDesigner designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                    };

                    // Example data source: a collection of Person objects.
                    ArrayList persons = new ArrayList
                    {
                        new Person { Name = "John Doe", Age = 30 },
                        new Person { Name = "Jane Smith", Age = 28 }
                    };

                    // Bind the data source to the smart marker name used in the template.
                    designer.SetDataSource("Persons", persons);

                    // Process all smart markers in the workbook.
                    designer.Process();

                    // Save the processed workbook to a memory stream and return its byte array.
                    using (MemoryStream resultStream = workbook.SaveToStream())
                    {
                        return resultStream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing template: {ex.Message}");
                return Array.Empty<byte>();
            }
        }
    }

    // Simple POCO class used as a data source for smart markers.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "result.xlsx";

            // Prevent FileNotFoundException for the template file.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found.");
                return;
            }

            try
            {
                byte[] templateBytes = File.ReadAllBytes(templatePath);
                byte[] resultBytes = SmartMarkerProcessor.ProcessTemplate(templateBytes);

                if (resultBytes.Length > 0)
                {
                    File.WriteAllBytes(outputPath, resultBytes);
                    Console.WriteLine($"Processed workbook saved to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("Processing failed; no output generated.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
