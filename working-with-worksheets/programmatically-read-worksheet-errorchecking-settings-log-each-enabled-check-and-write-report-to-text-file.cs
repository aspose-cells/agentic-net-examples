// Title: Read worksheet error‑checking options with Aspose.Cells for .NET and export enabled checks to a text report
// Common Searches: how to retrieve enabled error checking rules from an Excel worksheet using Aspose.Cells in C# | using reflection to access the ErrorCheckingOptions object in Aspose.Cells .NET | generate a text file that lists worksheet error‑checking settings with Aspose.Cells | enumerate boolean properties of ErrorCheckingOptions for each sheet in a workbook | log Excel worksheet validation options to a report using C# and Aspose.Cells
// Tags: Aspose.Cells read ErrorCheckingOptions via reflection | enumerate enabled error checks in Excel worksheet | export worksheet validation settings to text file C# | log worksheet error checking configuration Aspose.Cells | C# iterate worksheets error checking options

using System;
using System.IO;
using System.Reflection;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through all worksheets, uses reflection to obtain each sheet's ErrorCheckingOptions object, enumerates its boolean properties, records the names of checks that are enabled, and writes a formatted report of these settings to a text file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string reportPath = "ErrorCheckingReport.txt";

            // Ensure the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            StringBuilder reportBuilder = new StringBuilder();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                reportBuilder.AppendLine($"Worksheet: {sheet.Name}");

                // Try to obtain the ErrorCheckingOptions property via reflection
                PropertyInfo optionsProp = sheet.GetType().GetProperty(
                    "ErrorCheckingOptions",
                    BindingFlags.Public | BindingFlags.Instance);

                if (optionsProp != null)
                {
                    object options = optionsProp.GetValue(sheet);
                    if (options != null)
                    {
                        // Enumerate all boolean properties of the options object
                        PropertyInfo[] optionProperties = options.GetType().GetProperties(
                            BindingFlags.Public | BindingFlags.Instance);

                        foreach (PropertyInfo prop in optionProperties)
                        {
                            if (prop.PropertyType == typeof(bool))
                            {
                                bool isEnabled = (bool)prop.GetValue(options);
                                if (isEnabled)
                                {
                                    reportBuilder.AppendLine($"  Enabled check: {prop.Name}");
                                }
                            }
                        }
                    }
                    else
                    {
                        reportBuilder.AppendLine("  ErrorCheckingOptions instance is null.");
                    }
                }
                else
                {
                    reportBuilder.AppendLine("  ErrorCheckingOptions property not available in this Aspose.Cells version.");
                }

                reportBuilder.AppendLine();
            }

            // Write the report to a text file
            File.WriteAllText(reportPath, reportBuilder.ToString());
            Console.WriteLine($"Report generated: {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
