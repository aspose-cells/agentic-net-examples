// Title: Clone built‑in document properties from a template Excel workbook to a new workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a template .xlsx with Aspose.Cells, copies all built‑in document properties to a newly created workbook, and saves the result. | Generate a reusable method that transfers Excel metadata such as Author, Title, Keywords, and CreatedTime from one workbook to another using the Aspose.Cells API. | Provide a step‑by‑step example showing how to duplicate the built‑in properties of an Excel template into a fresh workbook in C#.
// Common Searches: Aspose.Cells copy built‑in document properties from one workbook to another C# | C# clone Excel file metadata using Aspose.Cells | How to transfer author and title properties between Excel workbooks with Aspose.Cells .NET | Programmatically duplicate Excel built‑in properties in a new workbook Aspose.Cells | Load template workbook and copy its properties to a new workbook Aspose.Cells example
// Tags: clone built-in document properties Aspose.Cells | copy Excel metadata C# Aspose.Cells | transfer workbook properties .NET | load template workbook Aspose.Cells | save workbook with cloned properties C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for a template Excel file, loads it with Aspose.Cells, creates an empty workbook, copies each built‑in document property (Author, Title, Subject, Keywords, Comments, LastSavedBy, Category, Manager, Company, CreatedTime, LastPrinted, LastSavedTime, Version) from the template to the new workbook, and saves the cloned workbook as ClonedWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "Template.xlsx";
            string outputPath = "ClonedWorkbook.xlsx";

            try
            {
                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file '{templatePath}' not found.");
                    return;
                }

                // Load the template workbook
                Workbook templateWorkbook = new Workbook(templatePath);

                // Create a new empty workbook
                Workbook newWorkbook = new Workbook();

                // Clone built‑in document properties from the template to the new workbook
                var srcProps = templateWorkbook.BuiltInDocumentProperties;
                var dstProps = newWorkbook.BuiltInDocumentProperties;

                dstProps.Author = srcProps.Author;
                dstProps.Title = srcProps.Title;
                dstProps.Subject = srcProps.Subject;
                dstProps.Keywords = srcProps.Keywords;
                dstProps.Comments = srcProps.Comments;
                dstProps.LastSavedBy = srcProps.LastSavedBy;
                dstProps.Category = srcProps.Category;
                dstProps.Manager = srcProps.Manager;
                dstProps.Company = srcProps.Company;
                dstProps.CreatedTime = srcProps.CreatedTime;
                dstProps.LastPrinted = srcProps.LastPrinted;
                dstProps.LastSavedTime = srcProps.LastSavedTime;
                dstProps.Version = srcProps.Version;

                // Save the new workbook with the cloned properties
                newWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
