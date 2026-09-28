// Title: C# console app to load an Excel file with Italian CultureInfo, add subtotal rows via reflection, and export to PDF using Aspose.Cells
// AI Prompts: Generate C# code that opens an .xlsx workbook with the Italian (it-IT) CultureInfo, inserts sum subtotals for the data range, and saves the result as a PDF using Aspose.Cells. | Demonstrate how to use .NET reflection to call the Worksheet.Subtotal method safely when the API may be absent in the current Aspose.Cells version. | Build a console program that validates the input Excel file, applies subtotals, gracefully handles missing Subtotal support, and writes the output as a PDF.
// Common Searches: aspnet load excel with it-it cultureinfo using aspose.cells | how to add subtotal rows in aspose.cells when Subtotal method is missing | export excel to pdf after adding subtotals with aspose.cells c# | invoke Worksheet.Subtotal via reflection in older Aspose.Cells releases | c# program to convert localized excel workbook to pdf with subtotals
// Tags: Italian CultureInfo loading Excel Aspose.Cells | reflection invoke Worksheet.Subtotal Aspose.Cells | add sum subtotals to worksheet Aspose.Cells | export workbook to PDF Aspose.Cells .NET | handle missing Subtotal API Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The console application checks for an input.xlsx file, loads it with Italian CultureInfo via LoadOptions, determines the data range, attempts to add sum subtotals using the Worksheet.Subtotal method through reflection (handling cases where the method is unavailable), and finally saves the workbook as output.pdf in PDF format.
class Program
{
    static void Main()
    {
        try
        {
            // Define Italian culture
            CultureInfo italianCulture = new CultureInfo("it-IT");

            // Verify that the input file exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook using the Italian culture settings
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = italianCulture
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the range of data (assumes a header row at row 0)
            int firstDataRow = 1; // data starts after header
            int lastDataRow = sheet.Cells.MaxDataRow;
            int lastDataColumn = sheet.Cells.MaxDataColumn;

            // Attempt to add subtotals using reflection (method may not exist in older versions)
            try
            {
                MethodInfo subtotalMethod = typeof(Worksheet).GetMethod(
                    "Subtotal",
                    BindingFlags.Instance | BindingFlags.Public,
                    null,
                    new Type[]
                    {
                        typeof(int), typeof(int), typeof(int), typeof(int),
                        typeof(object), typeof(int[]), typeof(int[]), typeof(bool)
                    },
                    null);

                if (subtotalMethod != null)
                {
                    // Resolve SubtotalType enum (if present)
                    Type subtotalEnum = typeof(Worksheet).Assembly.GetType("Aspose.Cells.SubtotalType");
                    object sumEnumValue = subtotalEnum != null
                        ? Enum.Parse(subtotalEnum, "Sum")
                        : null;

                    // Invoke the Subtotal method
                    subtotalMethod.Invoke(
                        sheet,
                        new object[]
                        {
                            firstDataRow,
                            0,
                            lastDataRow,
                            lastDataColumn,
                            sumEnumValue,
                            new int[] { 1 },
                            new int[] { 0 },
                            true
                        });
                }
                else
                {
                    Console.WriteLine("Subtotal method is not supported in the current Aspose.Cells version.");
                }
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                Console.WriteLine($"Error while adding subtotals: {tie.InnerException.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error while adding subtotals: {ex.Message}");
            }

            // Save the workbook as a PDF file
            const string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
