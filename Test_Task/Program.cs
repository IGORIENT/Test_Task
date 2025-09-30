using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Test_Task
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            #region Диалог выбора XLSX
            string xlsxPath;
            using (var ofd = new OpenFileDialog
            {
                Title = "Выберите Transmittal XLSX",
                Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(ofd.FileName))
                {
                    MessageBox.Show("Файл не выбран. Завершение.", "TransmittalBuilder",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                xlsxPath = ofd.FileName;
            }
            #endregion

            // 2) Папка проекта (..\..\.. от bin\Debug\net5.0-windows\)
            var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\.."));

            try
            {
                // 3) Читаем три колонки строго: Name, Path, Extension
                var items = ReadNamePathExtRows(xlsxPath); // List<(dir,name,ext)>

                // 4) Создаём структуру прямо в папке проекта
                foreach (var (dirRel, name, ext) in items)
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    // нормализуем расширение
                    var cleanExt = (ext ?? "").Trim();
                    if (!string.IsNullOrEmpty(cleanExt) && !cleanExt.StartsWith(".")) cleanExt = "." + cleanExt;

                    // имя файла: если уже содержит это расширение, не добавляем второй раз
                    var trimName = name.Trim();
                    if (!string.IsNullOrEmpty(cleanExt) && !trimName.EndsWith(cleanExt, StringComparison.OrdinalIgnoreCase))
                    {
                        trimName += cleanExt;
                    }

                    var relDir = (dirRel ?? "").Trim().Replace('\\', '/').Trim('/');

                    // абсолютные пути
                    var absDir = string.IsNullOrEmpty(relDir)
                        ? projectDir
                        : Path.Combine(projectDir, relDir.Replace('/', Path.DirectorySeparatorChar));

                    var absFile = Path.Combine(absDir, trimName);

                    // создаём
                    Directory.CreateDirectory(absDir);
                    if (!File.Exists(absFile))
                        File.WriteAllBytes(absFile, Array.Empty<byte>()); // пустой файл-заглушка
                }

                MessageBox.Show("Готово. Структура создана в папке проекта.",
                    "TransmittalBuilder", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "TransmittalBuilder",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== Чтение трёх колонок: Name, Path, Extension =====
        private static List<(string Dir, string Name, string Ext)> ReadNamePathExtRows(string xlsxPath)
        {
            using var wb = new XLWorkbook(xlsxPath);
            var ws = wb.Worksheets.First();

            var firstRow = ws.FirstRowUsed().RowNumber();
            var lastRow = ws.LastRowUsed().RowNumber();
            var firstCol = ws.FirstColumnUsed().ColumnNumber();
            var lastCol = ws.LastColumnUsed().ColumnNumber();

            // маппинг имён колонок (без регистра и пробелов по краям)
            var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = firstCol; c <= lastCol; c++)
            {
                var h = ws.Cell(firstRow, c).GetString().Trim();
                if (!string.IsNullOrEmpty(h)) headers[h] = c;
            }

            // ищем строго эти 3 названия;
            if (!headers.TryGetValue("Name", out var nameCol) ||
                !headers.TryGetValue("Path", out var pathCol) ||
                !headers.TryGetValue("Extension", out var extCol))
            {
                throw new InvalidOperationException("Ожидались колонки: Name, Path, Extension.");
            }

            var list = new List<(string Dir, string Name, string Ext)>();
            for (int r = firstRow + 1; r <= lastRow; r++)
            {
                var name = ws.Cell(r, nameCol).GetString();
                var dir = ws.Cell(r, pathCol).GetString();
                var ext = ws.Cell(r, extCol).GetString();

                // пропускаем полностью пустые строки
                if (string.IsNullOrWhiteSpace(name) ||
                    string.IsNullOrWhiteSpace(dir) ||
                    string.IsNullOrWhiteSpace(ext))
                    continue;

                list.Add((dir, name, ext));
            }

            return list;
        }
    }
}
