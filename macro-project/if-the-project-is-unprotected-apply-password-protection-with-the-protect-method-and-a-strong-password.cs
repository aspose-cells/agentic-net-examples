using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject class

namespace AsposeCellsExamples
{
    public class ProtectVbaProjectIfUnprotected
    {
        public static void Run()
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsm";
            string outputPath = "output_protected.xlsm";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project (may be null if workbook has no VBA)
                VbaProject vbaProject = workbook.VbaProject;

                // Protect the VBA project if it exists and is not already protected
                if (vbaProject != null && !vbaProject.IsProtected)
                {
                    // Protect the VBA project with a password (not locked for viewing)
                    vbaProject.Protect(false, "Str0ngP@ssw0rd!2026");
                }

                // Save the protected workbook
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console applications
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectVbaProjectIfUnprotected.Run();
        }
    }
}