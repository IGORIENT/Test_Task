using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Test_Task
{
    public class ExcelReader
    {
        public List<TrasmittalFile> Read(string xlsxPath)
        {
            using var workBook = new XLWorkbook(xlsxPath);
            var workSheets = workBook.Worksheets.First();

            var firstRow = workSheets.FirstRowUsed().RowNumber();
            var lastRow = workSheets.LastRowUsed().RowNumber();
            var firstCol = workSheets.FirstColumnUsed().ColumnNumber();
            var lastCol = workSheets.LastColumnUsed().ColumnNumber();

            // Заголовки
            var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = firstCol; c <= lastCol; c++)
            {
                var h = workSheets.Cell(firstRow, c).GetString().Trim();
                if (!string.IsNullOrEmpty(h)) headers[h] = c;
            }

            if (!headers.TryGetValue("Name", out int numberOfNameCol) ||
                !headers.TryGetValue("Path", out int numberOfPathCol) ||
                !headers.TryGetValue("Extension", out int numberOfExtCol))
                throw new InvalidOperationException("Ожидались колонки: Name, Path, Extension.");

            var list = new List<TrasmittalFile>();
            for (int r = firstRow + 1; r <= lastRow; r++)
            {
                var name = workSheets.Cell(r, numberOfNameCol).GetString();
                var dir = workSheets.Cell(r, numberOfPathCol).GetString()["Project\\".Length..]; 
                var ext = workSheets.Cell(r, numberOfExtCol).GetString();

                // пропускаем не полные строки
                if (string.IsNullOrWhiteSpace(name) &&
                    string.IsNullOrWhiteSpace(dir) &&
                    string.IsNullOrWhiteSpace(ext))
                    continue;

                list.Add(new TrasmittalFile(dir, name, ext));
            }

            return list;
        }
    }
}