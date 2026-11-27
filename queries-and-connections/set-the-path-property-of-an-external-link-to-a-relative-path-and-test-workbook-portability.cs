// Title: Set relative external link path and test workbook portability using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set external link DataSource to a relative path | How to make Excel workbooks portable with relative external links in Aspose.Cells | Calculate formulas after moving linked workbooks using Aspose.Cells .NET | Copy source and destination folders while preserving external links in Aspose.Cells | Resolve relative external links with AbsolutePath in Aspose.Cells C#
// Tags: relative external link DataSource Aspose.Cells C# | workbook portability Aspose.Cells .NET | calculate formulas with external links Aspose.Cells | copy Excel folders for linked workbooks Aspose.Cells | AbsolutePath property for resolving relative links Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkPathDemo
{
    // The example creates a source workbook (source.xlsx) with sample data, then builds a destination workbook (dest.xlsx) whose formulas reference the source file. It sets the external link's DataSource to a relative path ("..\Source\source.xlsx"), saves both workbooks, copies the Source and Destination directories to a new portable location, opens the copied destination workbook, assigns its AbsolutePath to the new folder, recalculates the formulas, and prints the retrieved values to confirm that the relative external link resolves correctly.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare folders
                // -----------------------------------------------------------------
                string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "DemoData");
                string sourceDir = Path.Combine(baseDir, "Source");
                string destDir = Path.Combine(baseDir, "Destination");
                string portableDir = Path.Combine(baseDir, "Portable");

                Directory.CreateDirectory(sourceDir);
                Directory.CreateDirectory(destDir);
                Directory.CreateDirectory(portableDir);

                // -----------------------------------------------------------------
                // 2. Create the source workbook (source.xlsx) with some data
                // -----------------------------------------------------------------
                string sourceFileName = "source.xlsx";
                string sourceFilePath = Path.Combine(sourceDir, sourceFileName);

                Workbook sourceWb = new Workbook();
                Worksheet sourceWs = sourceWb.Worksheets[0];
                sourceWs.Name = "DataSheet";
                sourceWs.Cells["A1"].PutValue("Hello from source!");
                sourceWs.Cells["A2"].PutValue(12345);
                sourceWb.Save(sourceFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Source workbook saved to: {sourceFilePath}");

                // -----------------------------------------------------------------
                // 3. Create the destination workbook (dest.xlsx) that links to source.xlsx
                // -----------------------------------------------------------------
                string destFileName = "dest.xlsx";
                string destFilePath = Path.Combine(destDir, destFileName);

                Workbook destWb = new Workbook();
                Worksheet destWs = destWb.Worksheets[0];
                destWs.Name = "LinkSheet";

                // Add formulas that reference the source workbook (file name only)
                destWs.Cells["A1"].Formula = $"='[source.xlsx]DataSheet'!A1";
                destWs.Cells["A2"].Formula = $"='[source.xlsx]DataSheet'!A2";

                // Set the external link to a relative path
                if (destWb.Worksheets.ExternalLinks.Count > 0)
                {
                    string relativePath = Path.Combine("..", "Source", sourceFileName);
                    destWb.Worksheets.ExternalLinks[0].DataSource = relativePath;
                    Console.WriteLine($"External link DataSource set to relative path: {relativePath}");
                }
                else
                {
                    Console.WriteLine("No external links were created – something went wrong.");
                    return;
                }

                destWb.Save(destFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Destination workbook saved to: {destFilePath}");

                // -----------------------------------------------------------------
                // 4. Simulate portability: copy both folders to a new location
                // -----------------------------------------------------------------
                string portableSourceDir = Path.Combine(portableDir, "Source");
                CopyDirectory(sourceDir, portableSourceDir);

                string portableDestDir = Path.Combine(portableDir, "Destination");
                CopyDirectory(destDir, portableDestDir);

                string portableDestPath = Path.Combine(portableDestDir, destFileName);
                Console.WriteLine("\n--- Portability Test ---");
                Console.WriteLine($"Opening destination workbook from portable location: {portableDestPath}");

                // -----------------------------------------------------------------
                // 5. Open the copied destination workbook and evaluate the formulas
                // -----------------------------------------------------------------
                if (!File.Exists(portableDestPath))
                {
                    Console.WriteLine("Portable destination workbook not found.");
                    return;
                }

                Workbook portableDestWb = new Workbook(portableDestPath);

                // Set the absolute path so that relative external links can be resolved
                portableDestWb.AbsolutePath = portableDestDir;

                // Calculate formulas (external links will be resolved using the relative path)
                portableDestWb.CalculateFormula();

                Worksheet portableDestWs = portableDestWb.Worksheets[0];
                Console.WriteLine($"Value of A1 (should be from source): {portableDestWs.Cells["A1"].StringValue}");
                Console.WriteLine($"Value of A2 (should be from source): {portableDestWs.Cells["A2"].StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Copies all files and sub‑directories from source to destination.
        static void CopyDirectory(string source, string destination)
        {
            try
            {
                Directory.CreateDirectory(destination);

                foreach (string filePath in Directory.GetFiles(source))
                {
                    string fileName = Path.GetFileName(filePath);
                    string destFile = Path.Combine(destination, fileName);
                    File.Copy(filePath, destFile, true);
                }

                foreach (string dirPath in Directory.GetDirectories(source))
                {
                    string dirName = Path.GetFileName(dirPath);
                    string destSubDir = Path.Combine(destination, dirName);
                    CopyDirectory(dirPath, destSubDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to copy directory '{source}' to '{destination}': {ex.Message}");
            }
        }
    }
}
