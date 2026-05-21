using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Folder that contains the Excel workbooks to be inspected
        string inputFolder = "InputWorkbooks";

        // Path of the CSV file that will hold the report
        string csvPath = "WorkbookReport.csv";

        // Retrieve all files in the folder and filter for Excel extensions
        string[] allFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        string[] excelFiles = Array.FindAll(allFiles, f =>
            f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

        // Create the CSV and write the header
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            writer.WriteLine("WorkbookName,HasMacro,IsDigitallySigned");

            // Process each workbook
            foreach (string filePath in excelFiles)
            {
                // Load the workbook using the provided constructor rule
                Workbook workbook = new Workbook(filePath);

                // Gather required properties
                bool hasMacro = workbook.HasMacro;               // macro assignment status
                bool isSigned = workbook.IsDigitallySigned;      // digital signature verification

                // Write a CSV line for the current workbook
                string line = $"{Path.GetFileName(filePath)},{hasMacro},{isSigned}";
                writer.WriteLine(line);

                // Release resources
                workbook.Dispose();
            }
        }

        // The CSV file "WorkbookReport.csv" now contains the desired information.
    }
}