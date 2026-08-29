// Title: Combine two Excel workbooks and export the merged workbook to PDF while preserving chart and image fidelity using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads two .xlsx files, merges them with Workbook.Combine, and saves the result as a PDF using Aspose.Cells. | Generate a .NET console application that verifies visual fidelity of charts and images after combining workbooks and exporting to PDF. | Create a C# example that checks for missing source files, creates empty workbooks if needed, combines them, and outputs a single PDF document.
// Common Searches: Aspose.Cells combine two workbooks and save as PDF C# | How to merge multiple Excel files and export to PDF with chart preservation using Aspose.Cells | C# code to load two .xlsx files, combine them, and generate a single PDF document | Export combined workbook to PDF while keeping images intact Aspose.Cells .NET
// Tags: Workbook.Combine merge workbooks Aspose.Cells | save workbook as PDF Aspose.Cells C# | preserve chart fidelity PDF Aspose.Cells | handle missing Excel file Aspose.Cells | export combined workbook to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportCombinedPdf
{
    // The program loads two Excel files (creating new workbooks if they are missing), merges the second workbook into the first using Workbook.Combine, and saves the combined workbook as a PDF, ensuring charts and images retain their visual fidelity.
    class Program
    {
        static void Main()
        {
            const string firstPath = "source1.xlsx";
            const string secondPath = "source2.xlsx";
            const string outputPdf = "CombinedOutput.pdf";

            Workbook firstWorkbook = null;
            Workbook secondWorkbook = null;

            try
            {
                // Load or create the first workbook.
                if (File.Exists(firstPath))
                {
                    firstWorkbook = new Workbook(firstPath);
                }
                else
                {
                    Console.WriteLine($"File '{firstPath}' not found. Creating a new workbook.");
                    firstWorkbook = new Workbook();
                }

                // Load or create the second workbook.
                if (File.Exists(secondPath))
                {
                    secondWorkbook = new Workbook(secondPath);
                }
                else
                {
                    Console.WriteLine($"File '{secondPath}' not found. Creating a new workbook.");
                    secondWorkbook = new Workbook();
                }

                // Combine the second workbook into the first one.
                firstWorkbook.Combine(secondWorkbook);

                // Save the combined workbook as PDF.
                firstWorkbook.Save(outputPdf, SaveFormat.Pdf);

                Console.WriteLine($"Combined workbook exported to PDF successfully: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up resources.
                firstWorkbook?.Dispose();
                secondWorkbook?.Dispose();
            }
        }
    }
}
