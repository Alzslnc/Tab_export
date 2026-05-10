using Autodesk.AutoCAD.DatabaseServices;
using BaseFunction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabExport.Data;
using static TabExport.SupportClass;

namespace TabExport
{
    internal class AcadTableExportClass
    {
        public static void Start()
        {
            if (!BaseGetObjectClass.TryGetIntFromUser(out int msl, Settings.Settings.Default.MaxStringLength, 1, null, "Ограничение числа символов в одной строке: ")) return;

            Settings.Settings.Default.MaxStringLength = msl;
            Settings.Settings.Save();

            if (!BaseGetObjectClass.TryGetobjectId(out ObjectId id, typeof(Table), "Выберите таблицу Autocad")) return;

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBObjectCollection collection = new DBObjectCollection();
                List<Line> lines = new List<Line>();
                Table table = null;
                try
                {
                    //получаем выбранную таблицу
                    table = tr.GetObject(id, OpenMode.ForRead, false, true).Clone() as Table;
                    if (table != null)
                    {

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                if (table.Cells[i, j].ContentTypes == CellContentTypes.Block) table.Cells[i, j].Value = "";
                            }
                        }

                        table.Explode(collection);

                        GetObjects(collection, out List<TextDataClass> texts, out lines);

                        //если нет текстов или линий то прекращаем
                        if (lines.Count == 0 || texts.Count == 0) return;

                        //формируем таблицу
                        TableStructureClass tableStructure = CreateTableStructure(texts, lines);

                        //если таблица не сформировалась то прекращаем
                        if (tableStructure.Columns.Count < 1) return;

                        //создаем эксель документ
                        ExcelClass.CreateExcelDocument.Create(tableStructure);
                    }
                }
                catch
                {
                }
                finally
                {
                    foreach (DBObject dBObject in collection) dBObject?.Dispose();
                    collection?.Dispose();
                    foreach (Curve curve in lines) curve?.Dispose();
                    table?.Dispose();
                }

                tr.Commit();
            }
        }
        public static void GetObjects(DBObjectCollection dBObjects, out List<TextDataClass> texts, out List<Line> lines)
        {
            texts = new List<TextDataClass>();
            lines = new List<Line>();

            foreach (DBObject bObject in dBObjects)
            {
                Entity entity = bObject as Entity;

                GetObject(entity, texts, lines);
            }
        }
        #region без расчленения
        internal static void Variant2()
        {
            if (!BaseGetObjectClass.TryGetIntFromUser(out int msl, Settings.Settings.Default.MaxStringLength, 1, null, "Ограничение числа символов в одной строке: ")) return;

            Settings.Settings.Default.MaxStringLength = msl;
            Settings.Settings.Save();

            if (!BaseGetObjectClass.TryGetobjectId(out ObjectId id, typeof(Table), "Выберите таблицу Autocad")) return;

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                try
                {
                    //получаем выбранную таблицу
                    Table table = tr.GetObject(id, OpenMode.ForRead, false, true) as Table;
                    if (table != null)
                    {
                        //формируем таблицу
                        TableStructureClass tableStructure = new TableStructureClass();
                        tableStructure.Cells = new DataCellClass[table.Rows.Count + 2, table.Columns.Count + 2];

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                DataCellClass dataCell = new DataCellClass { Row = i, Column = j, EndRow = i, EndColumn = j, Checked = true };

                                foreach (DataCellClass dataCellClass in tableStructure.Cells)
                                {
                                    if (dataCellClass == null) break;

                                    if (i >= dataCellClass.Row && i <= dataCellClass.EndRow && j >= dataCellClass.Column && j <= dataCellClass.EndColumn)
                                    {
                                        dataCell.Blocked = true;
                                        break;
                                    }
                                }

                                tableStructure.Cells[i, j] = dataCell;

                                if (dataCell.Blocked) continue;

                                if (table.Cells[i, j].IsMerged.HasValue && table.Cells[i, j].IsMerged.Value)
                                {
                                    CellRange cellReferences = table.Cells[i, j].GetMergeRange();

                                    dataCell.EndRow = cellReferences.BottomRow;
                                    dataCell.EndColumn = cellReferences.RightColumn;
                                }

                                if ((table.Cells[i, j].ContentTypes == CellContentTypes.Value || table.Cells[i, j].ContentTypes == CellContentTypes.Field) && table.Cells[i, j].Value != null)
                                {
                                    dataCell.Value = table.Cells[i, j].Value.ToString();
                                }

                            }
                        }

                        for (int i = table.Rows.Count - 1; i >= 0; i--)
                        {
                            for (int j = table.Columns.Count - 1; j >= 0; j--)
                            {
                                tableStructure.Cells[i + 1, j + 1] = tableStructure.Cells[i, j];
                                tableStructure.Cells[i, j].Column++;
                                tableStructure.Cells[i, j].Row++;
                                tableStructure.Cells[i, j].EndColumn++;
                                tableStructure.Cells[i, j].EndRow++;
                            }
                        }

                        for (int i = 0; i < table.Rows.Count + 2; i++)
                        {
                            tableStructure.Cells[i, 0] = new DataCellClass { Row = i, Column = 0, EndRow = i, EndColumn = 0, Checked = true, Blocked = true };
                        }
                      
                        for (int i = 0; i < table.Columns.Count + 2; i++)
                        {
                            tableStructure.Cells[0, i] = new DataCellClass { Row = 0, Column = i, EndRow = 0, EndColumn = i, Checked = true, Blocked = true };
                        }

                        int lastColumn = table.Columns.Count + 1;
                        for (int i = 0; i < table.Rows.Count + 2; i++)
                        {
                            tableStructure.Cells[i, lastColumn] = new DataCellClass { Row = i, Column = lastColumn, EndRow = i, EndColumn = lastColumn, Checked = true , Blocked = true};
                        }

                        int lastRow = table.Rows.Count + 1;
                        for (int i = 0; i < table.Columns.Count + 2; i++)
                        {
                            tableStructure.Cells[lastRow, i] = new DataCellClass { Row = lastRow, Column = i, EndRow = lastRow, EndColumn = i, Checked = true, Blocked = true };
                        }

                        //если таблица не сформировалась то прекращаем
                        if (tableStructure.Cells.Length < 2) return;

                        //создаем эксель документ
                        ExcelClass.CreateExcelDocument.Create(tableStructure);
                    }
                }
                catch 
                { 
                }
                finally
                {
                }

                tr.Commit();
            }
        }
        #endregion
    }
}
