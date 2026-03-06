using System;
using Aspose.Cells;

namespace AsposeCellsNonEnglishLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (must exist)
            string sourcePath = "sample.xlsx";

            // Create LoadOptions and set a non‑English UI language (e.g., German)
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguageCode = CountryCode.Germany; // notice for non‑English users

            // Load the workbook with the specified LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Example modification: write a note indicating the language used
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Loaded with LanguageCode = Germany");

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook loaded with LanguageCode=Germany and saved to '{outputPath}'.");
        }
    }
}