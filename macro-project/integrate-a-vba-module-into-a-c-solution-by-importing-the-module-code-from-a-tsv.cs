using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string tsvPath = Path.Combine(baseDir, "modules.tsv");

        if (!File.Exists(tsvPath))
        {
            string defaultContent = "Module1\tProcedural\tSub HelloWorld()\r\n    MsgBox \"Hello, World!\"\r\nEnd Sub";
            File.WriteAllText(tsvPath, defaultContent);
        }

        Workbook workbook = new Workbook();

        VbaProject vbaProject = workbook.VbaProject;

        foreach (string line in File.ReadLines(tsvPath))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split('\t');
            if (parts.Length < 3)
                continue;

            string moduleName = parts[0].Trim();
            string typeString = parts[1].Trim();
            string code = parts[2];

            if (!Enum.TryParse(typeString, true, out VbaModuleType moduleType))
                moduleType = VbaModuleType.Procedural;

            int index = vbaProject.Modules.Add(moduleType, moduleName);
            VbaModule module = vbaProject.Modules[index];
            module.Codes = code;
        }

        string outputPath = Path.Combine(baseDir, "output.xlsm");
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}