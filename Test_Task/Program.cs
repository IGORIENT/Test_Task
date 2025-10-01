using System;
using System.IO;
using System.Windows.Forms;

namespace Test_Task
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            #region  Диалог выбора XLSX
            string xlsxPath;
            using (var ofd = new OpenFileDialog
            {
                Title = "Выберите Transmittal XLSX",
                Filter = "Excel (*.xlsx)|*.xlsx",
                CheckFileExists = true,
                Multiselect = false
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(ofd.FileName))
                    return;

                xlsxPath = ofd.FileName;
            }
            #endregion


            // Папка проекта (..\..\.. от bin\Debug\net5.0-windows\)
            var projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\.."));

            try
            {
                // Читаем Excel
                var reader = new ExcelReader();
                var records = reader.Read(xlsxPath);

                // Строим структуру
                var builder = new StructureBuilder(projectDir);
                builder.Build(records);

                // Успешное сообщение
                MessageBox.Show(
                    $"Готово!\nФайлов создано: {records.Count}\nВ папке: {projectDir}",
                    "TransmittalBuilder",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message,
                    "TransmittalBuilder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
