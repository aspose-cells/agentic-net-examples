// Title: Apply a custom 0.00% number format to percentage cells in an Excel workbook using Aspose.Cells for .NET with Brazilian Portuguese culture
// Common Searches: how to apply a custom 0.00% number format to percent cells with Aspose.Cells in C# | set Brazilian Portuguese locale for Excel formatting using Aspose.Cells .NET | change existing percent styled cells to two‑decimal percentage format in Aspose.Cells | load an Excel file with optional password and format percentages in pt‑BR culture | iterate worksheets and update percent cell style with Aspose.Cells C#
// Tags: custom percentage format Aspose.Cells .NET | Brazilian Portuguese locale Excel formatting | apply style to percent cells C# | password protected workbook loading Aspose.Cells | set custom number format string Excel C# | culture‑specific number formatting Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example sets the thread culture to pt-BR, loads an Excel workbook (prompting for a password if required), scans all worksheets for cells whose style is marked as percent, assigns the custom number format "0.00%" to those cells, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Set Brazilian Portuguese culture for the current thread
        CultureInfo brazil = new CultureInfo("pt-BR");
        CultureInfo.CurrentCulture = brazil;
        CultureInfo.CurrentUICulture = brazil;

        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        // First attempt: load without password
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            // If loading fails, assume it may be password protected and request password
            Console.Write("The workbook may be password protected. Enter password (or leave empty to abort): ");
            string password = Console.ReadLine();

            if (string.IsNullOrEmpty(password))
            {
                Console.WriteLine("No password provided. Exiting.");
                return;
            }

            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to load workbook with password: {innerEx.Message}");
                return;
            }
        }

        // Ensure workbook was loaded
        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        // Iterate through all worksheets and apply a custom percentage format
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // Check if the cell is already formatted as a percentage
                    if (cell.GetStyle().IsPercent)
                    {
                        Style style = cell.GetStyle();
                        // Apply custom number format string for percentages
                        style.Custom = "0.00%";
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
