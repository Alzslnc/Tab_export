using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BaseFunction;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using TabExport.Data;
using Autodesk.AutoCAD.Geometry;

namespace TabExport
{
    
    public class TestClass
    {
        [CommandMethod("TableTest")]
        public void Start()
        {
            if (!BaseGetObjectClass.TryGetObjectsIds(out List<ObjectId> ids, new List<Type> { typeof(Line), typeof(Polyline), typeof(DBText), typeof(MText) })) return;

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {                 
                List<TextDataClass> texts = new List<TextDataClass>();               
                List<Line> lines = new List<Line>();

                List<double> textHeights = new List<double>();

                foreach (ObjectId id in ids)
                { 
                    Entity entity = tr.GetObject(id, OpenMode.ForRead, false, true) as Entity;

                    if (entity is DBText bText)
                    {
                        TextDataClass textData = new TextDataClass { TextHeight = bText.Height, VerticalValue = bText.Rotation > 1 && bText.Rotation < 2, Value = bText.TextString };
                        Polyline ex = bText.CreatePolyline();
                        textData.X = (ex.GeometricExtents.MaxPoint.X - ex.GeometricExtents.MinPoint.X) / 2;
                        textData.Y = (ex.GeometricExtents.MaxPoint.Y - ex.GeometricExtents.MinPoint.Y) / 2;
                        ex?.Dispose();
                        texts.Add(textData);
                    }
                    else if (entity is MText mText)
                    {
                        using (DBObjectCollection collection = new DBObjectCollection())
                        {
                            mText.Explode(collection);

                            foreach (DBObject dBObject in collection)
                            {
                                if (dBObject is DBText dBText)
                                {
                                    TextDataClass textData = new TextDataClass { TextHeight = dBText.Height, VerticalValue = dBText.Rotation > 1 && dBText.Rotation < 2, Value = dBText.TextString };
                                    Polyline ex = dBText.CreatePolyline();
                                    textData.X = (ex.GeometricExtents.MaxPoint.X - ex.GeometricExtents.MinPoint.X) / 2;
                                    textData.Y = (ex.GeometricExtents.MaxPoint.Y - ex.GeometricExtents.MinPoint.Y) / 2;
                                    ex?.Dispose();
                                    texts.Add(textData);
                                }
                                dBObject?.Dispose();
                            }
                        }                        
                    }
                    else if (entity is Line line)
                    {
                        lines.Add(new Line(line.StartPoint.Z0(), line.EndPoint.Z0()));
                    }
                    else if (entity is Polyline poly)
                    {
                        using (DBObjectCollection collection = new DBObjectCollection())
                        {
                            poly.Explode(collection);
                            foreach (DBObject dBObject in collection)
                            {
                                if (dBObject is Line l) lines.Add(new Line(l.StartPoint.Z0(), l.EndPoint.Z0()));
                                else dBObject?.Dispose();                                        
                            }
                        }
                    }
                }

                if (lines.Count == 0 || texts.Count == 0) return;

                TableStructureClass tableStructure = CreateTableStructure(texts, lines);



                //так как в списке могут быть результаты расчленения полилиний диспозим все
                foreach (Curve curve in lines) curve?.Dispose();

                tr.Commit();
            }
        }

        public TableStructureClass CreateTableStructure(List<TextDataClass> texts, List<Line> lines)
        {
            TableStructureClass result = new TableStructureClass();

            IEnumerable<Line> vertical = lines.Where(x => Math.Abs(x.StartPoint.Y - x.EndPoint.Y) > Math.Abs(x.StartPoint.X - x.EndPoint.X));
            IEnumerable<Line> horizontal = lines.Except(vertical);

            if (vertical.Count() < 2 || horizontal.Count() < 2) return result;

            //получаем область в которой объединяем координаты линий, то есть ширина или высота клетки должны быть больше чем половина средней высоты текстов
            double uniteRange = texts.Average(x => x.TextHeight) / 2;

            //получаем координаты по вертикали (сверху вниз, координаты строк)
            List<double> rowCoordinates = horizontal.Select(x => x.StartPoint.Y).Union(lines.Select(x => x.EndPoint.Y)).ToList();
            rowCoordinates.Sort();
            rowCoordinates.Reverse();
            rowCoordinates = GetRangeValues(rowCoordinates, uniteRange);

            //создаем список рядов (ряды считаются сверху вниз и координаты идут на убыль)
            //первай
            result.Rows.Add(new RangeClass() { Position = result.Rows.Count, StartPosition = double.MaxValue, EndPosition = rowCoordinates[0] });
            //промежуточные
            for (int i = 0; i < rowCoordinates.Count - 1; i++)
            {
                result.Rows.Add(new RangeClass() { Position = result.Rows.Count, StartPosition = rowCoordinates[i], EndPosition = rowCoordinates[i + 1] });
            }
            //последний
            result.Rows.Add(new RangeClass() { Position = result.Rows.Count, StartPosition = rowCoordinates[rowCoordinates.Count - 1], EndPosition = double.MinValue });

            //получаем координаты по горизонтали (координаты колонн)
            List<double> columnCoordinates = vertical.Select(x => x.StartPoint.X).Union(lines.Select(x => x.EndPoint.X)).ToList();
            columnCoordinates.Sort();
            columnCoordinates = GetRangeValues(columnCoordinates, uniteRange);

            //создаем список колонн
            //первая
            result.Columns.Add(new RangeClass() { Position = result.Columns.Count, StartPosition = double.MinValue, EndPosition = columnCoordinates[0] });
            //промежуточные
            for (int i = 0; i < columnCoordinates.Count - 1; i++)
            {
                result.Columns.Add(new RangeClass() { Position = result.Columns.Count, StartPosition = columnCoordinates[i], EndPosition = columnCoordinates[i + 1] });
            }
            //последняя
            result.Columns.Add(new RangeClass() { Position = result.Columns.Count, StartPosition = columnCoordinates[columnCoordinates.Count - 1], EndPosition = double.MaxValue });

            result.Cells = new DataCellClass[result.Rows.Count,result.Columns.Count];

            //создаем ячейки
            for (int i = 0; i < result.Rows.Count; i++)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    DataCellClass dataCell = new DataCellClass { Row = i, Column = j, EndRow = i, EndColumn = j };

                    IEnumerable<TextDataClass> textDatas = texts.Where(
                        x => x.Y >= result.Rows[i].EndPosition &&
                    x.Y < result.Rows[i].StartPosition &&
                    x.X >= result.Columns[j].StartPosition &&
                    x.X < result.Columns[j].EndPosition);

                    dataCell.TextDataClasses.AddRange(textDatas);

                    result.Cells[i, j] =  dataCell;
                }
            }

            //проверяем ячейки на объединение            
            for (int i = 1; i < result.Rows.Count - 1; i++)
            {
                for (int j = 1; j < result.Columns.Count - 1; j++)
                {
                    //получаем ячейку
                    DataCellClass dataCellClass = result.Cells[i, j];
                    //если она среди проверенных то пропускаем
                    if (dataCellClass.Checked) continue;
                    dataCellClass.Checked = true;
                    
                    //получаем координаты центра ячейки
                    double y = (result.Rows[i].StartPosition + result.Rows[i].EndPosition) / 2;
                    double x = (result.Columns[j].StartPosition + result.Columns[j].EndPosition) / 2;

                    //проходим по ячейкам вправо
                    for (int k = j + 1; k < result.Columns.Count - 1; k++)
                    {
                        //получаем координаты правой грани ячейки
                        Point3d point = new Point3d(result.Columns[j].StartPosition, y, 0);

                        //ищем ближайшую точку к грани
                        Point3d closest = vertical.MinBy(x => x.GetClosestPointTo(point, false)).GetClosestPointTo(point, false);

                        //если расстояние между точкой на грани и ближайшей точкой на линии меньше допуска то считаем что тут проходит граница, прерываем обход
                        if ((closest - point).Length < uniteRange) break;

                        //иначе объединяем ячейки
                        dataCellClass.EndColumn = k;

                        //объединенную ячейку объявляем проверенной
                        result.Cells[i, k].Checked = true;
                        //переносим тексты иэ добавленной ячейки к основной
                        dataCellClass.TextDataClasses.AddRange(result.Cells[i, k].TextDataClasses);
                        result.Cells[i, k].TextDataClasses.Clear();
                    }

                    //проходим по ячейкам вниз
                    for (int k = i + 1; k < result.Rows.Count - 1; k++)
                    {
                        //получаем координаты правой грани ячейки
                        Point3d point = new Point3d(x, result.Rows[k].StartPosition, 0);

                        //ищем ближайшую точку к грани
                        Point3d closest = horizontal.MinBy(x => x.GetClosestPointTo(point, false)).GetClosestPointTo(point, false);

                        //если расстояние между точкой на грани и ближайшей точкой на линии меньше допуска то считаем что тут проходит граница, прерываем обход
                        if ((closest - point).Length < uniteRange) break;

                        //иначе объединяем ячейки
                        dataCellClass.EndRow = k;

                        //проходим по всем ячейкам для объединения
                        for (int g = j; g <= dataCellClass.EndColumn; g++)
                        {
                            //объединенную ячейку объявляем проверенной
                            result.Cells[k, g].Checked = true;
                            //переносим тексты иэ добавленной ячейки к основной
                            dataCellClass.TextDataClasses.AddRange(result.Cells[k, g].TextDataClasses);
                            result.Cells[k, g].TextDataClasses.Clear();
                        }
                    }
                }
            }


            return result;

            //while (textDatas.Count() > 0)
            //{ 
            //    TextDataClass first = textDatas.MaxBy(x => x.Y);
            //    IEnumerable<TextDataClass> classes = textDatas.Where(x => Math.Abs(x.Y - first.Y) < first.TextHeight / 2);
            //    textDatas.;


            //}
        }
        private static List<double> GetRangeValues(List<double> coordinates, double uniteRange)
        {
            coordinates.ClearFromDoubles();
            List<double> result = new List<double>();
            while (coordinates.Count > 0)
            {
                List<double> unites = new List<double> { coordinates[coordinates.Count - 1] };            
                coordinates.RemoveAt(coordinates.Count - 1);

                for (int i = coordinates.Count - 1; i >= 0; i--)
                {
                    if (Math.Abs(coordinates[i] - unites.Last()) < uniteRange) result.Add(coordinates[i]);
                    else break;
                }

                result.Add(unites.Average());
            }
            return result;
        }
    }
}
