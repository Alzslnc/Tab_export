using System.Diagnostics;
using TabExport.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System;

namespace TabExport.ExcelClass
{
    internal static class CreateExcelDocument
    {
        public static void Create(TableStructureClass tableStructure)
        {
            ////создаем экземпляр прилоежения
            Excel.Application excApp;
            try
            {
                excApp = new Excel.Application();
            }
            catch
            {
                System.Windows.MessageBox.Show("Excel не найден", "Ошибка");
                return;
            }
            ////получаем процесс
            Process excelProc = Process.GetProcessesByName("EXCEL").Last();
            //создаем переменную страницы и книги
            Excel.Worksheet worksheet;
            Excel.Workbook book;
            //добавляем книгу приложению
            book = excApp.Workbooks.Add(); int worksheet_num = 1;
            //удаляем лишние страницы книга
            if (book.Worksheets.Count > 1)
            {
                for (int i = book.Worksheets.Count; i > 1; i--)
                {
                    worksheet = book.Worksheets.get_Item(i) as Excel.Worksheet;
                    worksheet.Delete();
                }
            }
            //добавляем страницу книге 
            if (book.Worksheets.Count >= worksheet_num)
            {
                worksheet = book.Worksheets.get_Item(worksheet_num) as Excel.Worksheet;
            }
            else
            {
                book.Worksheets.Add(After: book.Worksheets[worksheet_num - 1]);
                worksheet = book.Worksheets.get_Item(worksheet_num) as Excel.Worksheet;
            }
            //нумеруем страницу
            worksheet.Name = worksheet_num.ToString();

            //получаем область таблицы
            Excel.Range table = worksheet.get_Range((Excel.Range)worksheet.Cells[2, 2], (Excel.Range)worksheet.Cells[tableStructure.Cells.GetLength(0) - 1, tableStructure.Cells.GetLength(1) - 1]);

            table.Rows.RowHeight = 50;
            table.Columns.ColumnWidth = 100;

            foreach (DataCellClass dataCell in tableStructure.Cells)
            {
                //если это часть объединенной ячейки - пропускаем
                if (dataCell.Blocked) continue;

                //добавляем единицу к индексам строк/столбцов так как в экселе они идут с 1 а не 0
                dataCell.Column++;
                dataCell.Row++;
                dataCell.EndColumn++;
                dataCell.EndRow++;

                //проверяем, является ли ячейка объединенной и объединяем если требуется
                if (dataCell.EndColumn > dataCell.Column || dataCell.EndRow > dataCell.Row)
                { 
                    Excel.Range merged = worksheet.get_Range((Excel.Range)worksheet.Cells[dataCell.Row, dataCell.Column], (Excel.Range)worksheet.Cells[dataCell.EndRow, dataCell.EndColumn]);
                    merged.Merge();        
                }

                //если текста нет то пропускаем
                if (string.IsNullOrEmpty(dataCell.Value)) continue;

                //добавляем текст
                worksheet.Cells[dataCell.Row, dataCell.Column].Value = dataCell.Value;
                worksheet.Cells[dataCell.Row, dataCell.Column].NumberFormat = Format(dataCell.Value);

                //устанавливаем вертикальность
                if (dataCell.VerticalValue)
                {
                    Excel.Range range = worksheet.get_Range((Excel.Range)worksheet.Cells[dataCell.Row, dataCell.Column], (Excel.Range)worksheet.Cells[dataCell.EndRow, dataCell.EndColumn]);
                    range.Orientation = 90;
                }
            }

            //устанавливаем границы и ширину ячеек
            try
            {
                
                Excel.Borders borders = table.Borders;
               
                table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                table.Borders.Weight = Excel.XlBorderWeight.xlThin;
                table.Borders.Color =  Color.Black;

                table = worksheet.get_Range((Excel.Range)worksheet.Cells[1, 1], (Excel.Range)worksheet.Cells[tableStructure.Cells.GetLength(0), tableStructure.Cells.GetLength(1)]);  
              
                table.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; 
                table.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                table.Columns.AutoFit();
                table.Rows.AutoFit();

            }
            catch { }

            excApp.UserControl = true;
            excApp.Visible = true;
        }

        public static string Format(string value)
        {
            value = value.Trim();            
            if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                string result = "0";
                string[] strings = value.Split(new string[] { ",", "." }, StringSplitOptions.RemoveEmptyEntries);
                
                if (strings.Length > 1)
                {
                    result += ",";
                    for (int i = 0; i < strings[1].Length; i++) result += "0";
                }
                
                return result;
            }
            else return "@";
        }
    }
}
