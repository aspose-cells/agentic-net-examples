// Title: Generate a C# console report of each Excel workbook’s CheckCompatibility flag and custom ribbon status using Aspose.Cells
// AI Prompts: Write a C# console program that iterates over a list of Excel file paths, loads each workbook with Aspose.Cells, reads Workbook.Settings.CheckCompatibility and Workbook.RibbonXml, and prints a formatted table to the console. | Enhance the program to also display the workbook’s file format (XLSX, XLS, XLSM) alongside the compatibility flag and ribbon status. | Add functionality that writes the generated summary (file path, compatibility flag, ribbon status, file format) to a CSV file for further analysis.
// Common Searches: aspnet cells read checkcompatibility property for multiple workbooks | c# detect custom ribbonxml in Excel files using Aspose.Cells | how to list workbook settings in a console app with Aspose.Cells | export Aspose.Cells workbook summary to csv in C# | handle password‑protected Excel files when reading settings with Aspose.Cells
// Tags: read workbook checkcompatibility aspose.cells | detect custom ribbonxml aspose.cells | c# console generate workbook settings report | process multiple excel files aspose.cells | export workbook summary to csv c# | handle password protected workbooks aspose.cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookSummaryReport
{
    // The program loops through a predefined collection of Excel file paths, loads each workbook with Aspose.Cells, extracts the CheckCompatibility flag from Workbook.Settings and determines whether a custom RibbonXml is present. It prints a neatly aligned table showing the file path, compatibility status, and ribbon status, while gracefully handling missing files, password‑protected workbooks, and other load errors.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            var workbookPaths = new List<string>
            {
                "Sample1.xlsx",
                "Sample2.xls",
                "Sample3.xlsm"
                // Add more paths as needed
            };

            Console.WriteLine("Processed Workbook Summary");
            Console.WriteLine("---------------------------");
            Console.WriteLine("{0,-30} {1,-15} {2,-12}", "File Path", "CheckCompatibility", "RibbonStatus");
            Console.WriteLine(new string('-', 60));

            foreach (var path in workbookPaths)
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(path))
                {
                    Console.WriteLine("{0,-30} {1,-15} {2,-12}", path, "N/A", "File not found");
                    continue;
                }

                try
                {
                    // Load the workbook; if the file is password‑protected an exception will be thrown
                    using (Workbook workbook = new Workbook(path))
                    {
                        // Retrieve the compatibility setting from WorkbookSettings
                        bool checkCompatibility = workbook.Settings.CheckCompatibility;

                        // Determine ribbon status based on the RibbonXml property
                        string ribbonStatus = string.IsNullOrEmpty(workbook.RibbonXml) ? "None" : "Custom";

                        // Output the information for the current workbook
                        Console.WriteLine("{0,-30} {1,-15} {2,-12}",
                            path,
                            checkCompatibility,
                            ribbonStatus);
                    }
                }
                // Specific handling for password‑protected files (message contains "Password")
                catch (CellsException ex) when (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("{0,-30} {1,-15} {2,-12}", path, "N/A", "Password protected");
                }
                catch (Exception ex)
                {
                    // General error handling for any other issues (corrupt file, unsupported format, etc.)
                    Console.WriteLine("{0,-30} {1,-15} {2,-12}", path, "Error", ex.Message);
                }
            }

            Console.WriteLine("\nReport generation completed.");
        }
    }
}
