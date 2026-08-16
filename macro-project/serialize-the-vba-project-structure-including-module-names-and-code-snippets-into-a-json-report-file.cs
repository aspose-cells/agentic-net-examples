// Title: Export VBA Project from .xlsm to JSON with Aspose.Cells for .NET
// Description: Loads a macro‑enabled Excel workbook using Aspose.Cells, extracts the VBA project metadata, modules (name, type, source code) and references, and writes a formatted UTF‑8 JSON report. Includes signing, validation and protection flags for compliance and migration scenarios.
// Keywords: Aspose.Cells VBA export | C# serialize .xlsm VBA project | extract VBA modules JSON | Excel macro project metadata | VBA references Aspose.Cells | code documentation JSON | compliance audit VBA | migration of VBA code | macro‑enabled workbook analysis | System.Text.Json Excel VBA
// Common Searches: how to export VBA modules from xlsm using Aspose.Cells | C# generate JSON report of Excel VBA project | extract VBA references from macro‑enabled workbook | serialize VBA project metadata to JSON | Aspose.Cells read VBA code in .xlsm
// Developer Intent: Produce a JSON file that details the VBA project, its modules, and references from a macro‑enabled Excel workbook.
// Use Cases: Document and review macro code by exporting module names and source to JSON. | Run automated compliance checks on VBA signing, validation, and protection flags. | Facilitate migration of VBA logic to other platforms by extracting code snippets and reference data. | Create searchable inventories of VBA projects across multiple workbooks.
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsm file and outputs a JSON file containing each VBA module's name, type, and source code. | Provide a C# example that enumerates all VBA project references in an Excel workbook and saves them in a structured JSON report. | Generate a C# snippet that captures VBA project metadata (name, encoding, signing status, protection) and serializes it to an indented JSON document.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaProjectJsonExport
{
    // Classes that represent the JSON structure
    // Loads a macro‑enabled Excel workbook using Aspose.Cells, extracts the VBA project metadata, modules (name, type, source code) and references, and writes a formatted UTF‑8 JSON report. Includes signing, validation and protection flags for compliance and migration scenarios.
    public class VbaProjectReport
    {
        public string Name { get; set; } = string.Empty;
        public string Encoding { get; set; } = string.Empty;
        public bool IsSigned { get; set; }
        public bool IsValidSigned { get; set; }
        public bool IsProtected { get; set; }
        public List<VbaModuleInfo> Modules { get; set; } = new List<VbaModuleInfo>();
        public List<VbaReferenceInfo> References { get; set; } = new List<VbaReferenceInfo>();
    }

    public class VbaModuleInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Codes { get; set; } = string.Empty;
    }

    public class VbaReferenceInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file (must be macro-enabled, e.g., .xlsm)
            string inputPath = "InputWorkbook.xlsm";
            // Output JSON report file
            string outputPath = "VbaProjectReport.json";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Build the report object
                VbaProjectReport report = new VbaProjectReport
                {
                    Name = vbaProject.Name,
                    Encoding = vbaProject.Encoding?.WebName ?? string.Empty,
                    IsSigned = vbaProject.IsSigned,
                    IsValidSigned = vbaProject.IsValidSigned,
                    IsProtected = vbaProject.IsProtected
                };

                // Serialize modules
                foreach (VbaModule module in vbaProject.Modules)
                {
                    try
                    {
                        report.Modules.Add(new VbaModuleInfo
                        {
                            Name = module.Name,
                            Type = module.Type.ToString(),
                            Codes = module.Codes ?? string.Empty
                        });
                    }
                    catch (Exception modEx)
                    {
                        Console.WriteLine($"Error processing module '{module?.Name}': {modEx.Message}");
                    }
                }

                // Serialize references (Value property may not be available in some versions)
                foreach (VbaProjectReference reference in vbaProject.References)
                {
                    try
                    {
                        report.References.Add(new VbaReferenceInfo
                        {
                            Name = reference.Name,
                            Type = reference.Type.ToString(),
                            // Use empty string if the Value property is unavailable
                            Value = string.Empty
                        });
                    }
                    catch (Exception refEx)
                    {
                        Console.WriteLine($"Error processing reference '{reference?.Name}': {refEx.Message}");
                    }
                }

                // Convert the report to JSON (using System.Text.Json)
                JsonSerializerOptions options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                string json = JsonSerializer.Serialize(report, options);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to file
                File.WriteAllText(outputPath, json, Encoding.UTF8);

                Console.WriteLine($"VBA project structure has been serialized to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
