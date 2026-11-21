// Title: Populate an Excel template from a C# object and obtain a MemoryStream using Aspose.Cells
// Common Searches: C# Aspose.Cells populate Excel template with POCO and retrieve as MemoryStream | How to map C# object fields to named cells in an Excel workbook using Aspose.Cells | Saving an Aspose.Cells workbook to a MemoryStream instead of a file | Loading an Excel template and filling it with data from a C# object using Aspose.Cells | Best way to handle missing template file when using Aspose.Cells in .NET
// Tags: populate workbook using object data Aspose.Cells | load Excel template file C# Aspose.Cells | map object properties to named cells Aspose.Cells | export workbook to MemoryStream Aspose.Cells | template file not found handling Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace MyApp
{
    // The PopulateWorkbook method checks the input object and template path, loads the specified Excel template with Aspose.Cells, iterates over the object's public properties, writes each value into a cell whose name matches the property name, saves the modified workbook to a MemoryStream in XLSX format, resets the stream position, and returns the stream, while handling missing files and load errors.
    public class WorkbookHelper
    {
        /// <param name="data">An object whose public properties will be written to the workbook.</param>
        /// <param name="templatePath">Full path to the Excel template file.</param>
        /// <returns>A MemoryStream containing the populated workbook.</returns>
        public MemoryStream PopulateWorkbook(object data, string templatePath)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (string.IsNullOrWhiteSpace(templatePath)) throw new ArgumentException("Template path is required.", nameof(templatePath));

            // Ensure the template file exists before attempting to load it.
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Template file not found.", templatePath);

            Workbook workbook;
            try
            {
                // Load the template workbook.
                workbook = new Workbook(templatePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to load workbook template.", ex);
            }

            // Use the first worksheet (adjust if needed).
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate over all public instance properties of the data object.
            PropertyInfo[] properties = data.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in properties)
            {
                try
                {
                    // Attempt to retrieve a cell with the same name as the property.
                    Cell cell = cells[prop.Name];
                    if (cell != null)
                    {
                        // Write the property value into the cell.
                        cell.PutValue(prop.GetValue(data, null));
                    }
                }
                catch
                {
                    // If the cell name does not exist, skip this property.
                }
            }

            // Save the populated workbook into a memory stream.
            MemoryStream stream = new MemoryStream();
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0; // Reset stream position for reading.

            return stream;
        }
    }

    // Sample data class used in the demo.
    public class SampleData
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public double Salary { get; set; }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string templatePath = "template.xlsx";

                // Verify that the template file exists.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file '{templatePath}' not found.");
                    return;
                }

                // Create sample data to populate the workbook.
                var data = new SampleData
                {
                    Name = "John Doe",
                    Age = 30,
                    Salary = 12345.67
                };

                var helper = new WorkbookHelper();

                // Populate the workbook and obtain the result as a MemoryStream.
                using (MemoryStream ms = helper.PopulateWorkbook(data, templatePath))
                {
                    string outputPath = "output.xlsx";

                    // Write the stream to a physical file.
                    using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        ms.CopyTo(file);
                    }

                    Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
