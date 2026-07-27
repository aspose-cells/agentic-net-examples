using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace WorkbookMetadataUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the source workbook and the result workbook
            string sourcePath = "input.xlsx";
            string resultPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook from the source file
                Workbook workbook = new Workbook(sourcePath);

                // Determine if a worksheet named "Metadata" already exists
                Worksheet metadataSheet = workbook.Worksheets["Metadata"];
                if (metadataSheet == null)
                {
                    // Add a new worksheet and set its name to "Metadata"
                    int newIndex = workbook.Worksheets.Add();
                    metadataSheet = workbook.Worksheets[newIndex];
                    metadataSheet.Name = "Metadata";
                }

                // Write header titles
                metadataSheet.Cells["A1"].PutValue("Property Name");
                metadataSheet.Cells["B1"].PutValue("Property Value");

                // Start writing data from the second row (index 1)
                int currentRow = 1;

                // Iterate through all custom document properties of the workbook
                if (workbook.CustomDocumentProperties != null)
                {
                    foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
                    {
                        // Write property name
                        metadataSheet.Cells[currentRow, 0].PutValue(prop.Name);

                        // Write property value (convert to string safely)
                        string valueText = prop.Value != null ? prop.Value.ToString() : string.Empty;
                        metadataSheet.Cells[currentRow, 1].PutValue(valueText);

                        currentRow++;
                    }
                }

                // Ensure the directory for the result file exists
                string resultDir = Path.GetDirectoryName(resultPath);
                if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
                {
                    Directory.CreateDirectory(resultDir);
                }

                // Save the modified workbook to the result file
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to {resultPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}