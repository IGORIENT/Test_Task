using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Test_Task
{
    public class StructureBuilder
    {
        private readonly string _projectDir;

        public StructureBuilder(string projectDir)
        {
            _projectDir = projectDir;
        }

        public void Build(IEnumerable<TrasmittalFile> records)
        {
            foreach (var rec in records)
            {
                var Ext = (rec.Ext).Trim();
                if (!Ext.StartsWith("."))
                    Ext = "." + Ext;

                // имя файла
                var Name =rec.Name + Ext;
                

                // папка
                var relDir = rec.Dir.Trim().Replace('\\', '/').Trim('/');
                var absDir = Path.Combine(_projectDir, relDir.Replace('/', Path.DirectorySeparatorChar));

                var absFile = Path.Combine(absDir, Name);

                Directory.CreateDirectory(absDir);
                if (!File.Exists(absFile))
                    File.WriteAllBytes(absFile, Array.Empty<byte>());
            }
        }
    }
}
