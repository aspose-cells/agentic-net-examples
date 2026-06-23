using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchFreezePanesDemo
{
    class Program
    {
        static void Main()
        {
            // List of source workbook file paths (10 workbooks)
            List<string> sourceFiles = new List<string>
            {
                "input1.xlsx",
                "input2.xlsx",
                "input3.xlsx",
                "input4.xlsx",
                "input5.xlsx",
                "input6.xlsx",
                "input7.xlsx",
                "input8.xlsx",
                "input9.xlsx",
                "input10.xlsx"
            };

            // Corresponding output file paths
            List<string> outputFiles = new List<string>
            {
                "output1.xlsx",
                "output2.xlsx",
                "output3.xlsx",
                "output4.xlsx",
                "output5.xlsx",
                "output6.xlsx",
                "output7.xlsx",
                "output8.xlsx",
                "output9.xlsx",
                "output10.xlsx"
            };

            // Freeze configuration: freeze panes at cell "C3" with 2 frozen rows and 2 frozen columns
            string freezeCell = "C3";
            int frozenRows = 2;
            int frozenColumns = 2;

            // Process each workbook
            for (int i = 0; i < sourceFiles.Count; i++)
            {
                try
                {
                    // Verify source file exists
                    if (!File.Exists(sourceFiles[i]))
                    {
                        Console.WriteLine($"Source file not found: {sourceFiles[i]}");
                        continue;
                    }

                    // Load workbook, apply freeze panes, and save
                    using (Workbook workbook = new Workbook(sourceFiles[i]))
                    {
                        Worksheet sheet = workbook.Worksheets[0];
                        sheet.FreezePanes(freezeCell, frozenRows, frozenColumns);
                        workbook.Save(outputFiles[i]);
                    }

                    Console.WriteLine($"Processed: {sourceFiles[i]} -> {outputFiles[i]}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourceFiles[i]}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch freeze panes operation completed.");
        }
    }
}