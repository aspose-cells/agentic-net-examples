using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeReport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook (replace with actual file path)
                string inputFile = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputFile)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Prepare a StringWriter to collect the report
                using StringWriter report = new StringWriter();

                // Header for the report
                report.WriteLine("Named Ranges Report");
                report.WriteLine("====================");
                report.WriteLine();

                // Access the collection of defined names in the workbook
                NameCollection names = workbook.Worksheets.Names;

                // Iterate through each defined name
                foreach (Name name in names)
                {
                    // Name text (e.g., "MyRange")
                    string nameText = name.Text;

                    // The formula that the name refers to (starts with '=')
                    string refersTo = name.RefersTo;

                    // Write basic information
                    report.WriteLine($"Name: {nameText}");
                    report.WriteLine($"RefersTo: {refersTo}");

                    // Attempt to retrieve the actual ranges (if the name refers to a range)
                    try
                    {
                        // Get all ranges referred by this name (recalculates if needed)
                        AsposeRange[] ranges = name.GetRanges();

                        if (ranges != null && ranges.Length > 0)
                        {
                            report.WriteLine("Associated Ranges:");
                            foreach (AsposeRange rng in ranges)
                            {
                                // Each range address (e.g., "Sheet1!A1:B3")
                                report.WriteLine($"  - {rng.Address}");
                            }
                        }
                        else
                        {
                            report.WriteLine("Associated Ranges: None or external reference");
                        }
                    }
                    catch (Exception ex)
                    {
                        // In case GetRanges throws (e.g., unsupported reference)
                        report.WriteLine($"Error retrieving ranges: {ex.Message}");
                    }

                    report.WriteLine(); // Blank line between entries
                }

                // Output the report to console
                Console.WriteLine(report.ToString());

                // Optionally, save the report to a text file
                string reportPath = "NamedRangesReport.txt";
                File.WriteAllText(reportPath, report.ToString());
                Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");

                // Save the workbook (if any modifications were made)
                string outputFile = "output.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions to prevent the program from crashing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}