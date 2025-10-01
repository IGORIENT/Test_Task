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
                // расширение с точкой в начале
                var Ext = rec.Ext.StartsWith(".") ? rec.Ext : "." + rec.Ext;

                // имя файла + расширение
                var Name =rec.Name + Ext;

                // абсолютный путь к папке
                var absDir = Path.Combine(_projectDir, rec.Dir.Replace('\\', Path.DirectorySeparatorChar).Trim(Path.DirectorySeparatorChar));

                // полный путь к файлу
                var absFile = Path.Combine(absDir, Name);

                Directory.CreateDirectory(absDir); //создаю дирректорию
                
                if (!File.Exists(absFile))
                    File.WriteAllBytes(absFile, Array.Empty<byte>()); //создаю пустой файл-заглушку
            }
        }
    }
}
