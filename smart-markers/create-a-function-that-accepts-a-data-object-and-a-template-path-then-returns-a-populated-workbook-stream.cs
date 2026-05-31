using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookUtility
{
    public static class WorkbookHelper
    {
        /// <summary>
        /// Populates a workbook template with the provided data object and returns the result as a memory stream.
        /// </summary>
        /// <param name="data">The data source to bind to the smart markers in the template.</param>
        /// <param name="templatePath">Full path to the Excel template file containing smart markers.</param>
        /// <returns>A MemoryStream containing the populated workbook in XLSX format.</returns>
        public static MemoryStream PopulateWorkbook(object data, string templatePath)
        {
            // Ensure the template file exists to avoid FileNotFoundException.
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            try
            {
                // Load the template workbook.
                var workbook = new Workbook(templatePath);

                // Associate the workbook with a designer.
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind the data object to the smart marker name "Data".
                designer.SetDataSource("Data", data);
                designer.Process();

                // Save the processed workbook into a memory stream (XLSX format).
                var resultStream = new MemoryStream();
                workbook.Save(resultStream, SaveFormat.Xlsx);
                resultStream.Position = 0; // Reset for reading.

                return resultStream;
            }
            catch (Exception ex)
            {
                // Wrap any exception for clearer diagnostics.
                throw new InvalidOperationException("Failed to populate workbook.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage of PopulateWorkbook.
            string templatePath = "Template.xlsx";

            // Sample data matching the smart markers in the template.
            var data = new[]
            {
                new { Name = "Alice", Age = 30 },
                new { Name = "Bob", Age = 25 }
            };

            try
            {
                using var populatedStream = WorkbookHelper.PopulateWorkbook(data, templatePath);

                // Write the result to a file for verification.
                string outputPath = "Result.xlsx";
                using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                populatedStream.CopyTo(fileStream);

                Console.WriteLine($"Workbook generated successfully: {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}