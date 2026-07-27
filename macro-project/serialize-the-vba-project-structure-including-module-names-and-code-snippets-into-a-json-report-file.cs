using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text.Json;

namespace VbaProjectSerialization
{
    // Simple DTO for a VBA module
    public class VbaModuleInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Codes { get; set; }
    }

    // Simple DTO for a VBA project
    public class VbaProjectInfo
    {
        public string Name { get; set; }
        public string Encoding { get; set; }
        public bool IsSigned { get; set; }
        public bool IsProtected { get; set; }
        public bool IsLockedForViewing { get; set; }
        public List<VbaModuleInfo> Modules { get; set; } = new List<VbaModuleInfo>();
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook (macro‑enabled format will be used later)
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // ---------------------------------------------------------------
            // 2. Access the VBA project and set some basic properties
            // ---------------------------------------------------------------
            VbaProject vbaProject = workbook.VbaProject;
            vbaProject.Name = "SampleVbaProject";
            vbaProject.Encoding = Encoding.UTF8;

            // ---------------------------------------------------------------
            // 3. Add a couple of VBA modules and fill them with code
            // ---------------------------------------------------------------
            int idx1 = vbaProject.Modules.Add(VbaModuleType.Class, "ClassModule1");
            VbaModule module1 = vbaProject.Modules[idx1];
            module1.Codes = "Public Sub HelloWorld()\r\n    MsgBox \"Hello from ClassModule1\"\r\nEnd Sub";

            int idx2 = vbaProject.Modules.Add(VbaModuleType.Procedural, "StandardModule1");
            VbaModule module2 = vbaProject.Modules[idx2];
            module2.Codes = "Sub ShowMessage()\r\n    MsgBox \"Message from StandardModule1\"\r\nEnd Sub";

            // ---------------------------------------------------------------
            // 4. Build the DTO that will be serialized to JSON
            // ---------------------------------------------------------------
            VbaProjectInfo projectInfo = new VbaProjectInfo
            {
                Name = vbaProject.Name,
                Encoding = vbaProject.Encoding?.WebName,
                IsSigned = vbaProject.IsSigned,
                IsProtected = vbaProject.IsProtected,
                IsLockedForViewing = vbaProject.IslockedForViewing
            };

            foreach (VbaModule mod in vbaProject.Modules)
            {
                projectInfo.Modules.Add(new VbaModuleInfo
                {
                    Name = mod.Name,
                    Type = mod.Type.ToString(),
                    Codes = mod.Codes
                });
            }

            // ---------------------------------------------------------------
            // 5. Serialize the DTO to a formatted JSON string
            // ---------------------------------------------------------------
            string json = JsonSerializer.Serialize(projectInfo, new JsonSerializerOptions { WriteIndented = true });

            // ---------------------------------------------------------------
            // 6. Save the JSON report to disk
            // ---------------------------------------------------------------
            string jsonPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "VbaProjectReport.json");
            File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // ---------------------------------------------------------------
            // 7. Save the workbook as a macro‑enabled file (XLSM)
            // ---------------------------------------------------------------
            string workbookPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "SampleWorkbook.xlsm");
            workbook.Save(workbookPath, SaveFormat.Xlsm);

            Console.WriteLine($"VBA project JSON report saved to: {jsonPath}");
            Console.WriteLine($"Workbook saved to: {workbookPath}");
        }
    }
}