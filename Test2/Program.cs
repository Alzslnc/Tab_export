using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using winForm = System.Windows.Forms;

namespace Tab_export
{
    /// <summary>
    /// команды:
    /// 1 перенос взорванных (из линий и текстов) таблиц в эксель
    /// </summary>
    public class Program
    {
        //запуск первым вариантом команды
        [CommandMethod("Export_tab", CommandFlags.UsePickSet)]
        public void asdasdasdas()
        {
            Test91991();
        }
        readonly Function function = new Function();

        //запуск вторым вариантом команды
        [CommandMethod("bad_tab", CommandFlags.UsePickSet)]
        public void Test91991()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            //запуск меню изменения параметров если надо
            Editor ed = doc.Editor;
            PromptSelectionResult result = ed.SelectImplied();
            SelectionSet coll = result.Value;
            List<ObjectId> objectIds = new List<ObjectId>();
            if (coll != null)
            {
                objectIds.AddRange(coll.GetObjectIds());
            }
            if (objectIds.Count == 0)
            {
                //задаем параметры выбора ключевого слова
                PromptKeywordOptions vib = new PromptKeywordOptions("");
                //текст при выборе ключевого слова
                vib.Message = "\nВыбрать параметры?: ";
                //первый вариант выбор ключевого слова
                vib.Keywords.Add("Нет");
                //второй вариант выбора ключевого слова
                vib.Keywords.Add("Да");
                //нельзя ничего не выбрать
                vib.AllowNone = false;
                //по умолчанию выбор нет
                vib.Keywords.Default = "Нет";
                //запрос результата выбора ключевого слова
                PromptResult vibRes = doc.Editor.GetKeywords(vib);
                //если ключевое слово "Да" то запускаем форму для выбора параметров
                if (vibRes.StringResult == "Да")
                {
                    //создаем форму и отправляем в нее значения по умолчанию
                    bad_tab_form form = new bad_tab_form();
                    //получаем ответ от формы
                    winForm.DialogResult dialRes = Application.ShowModalDialog(form);
                    form.Dispose();
                }
                //список элементов для множественного выбора
                List<string> objTypes = new List<string>();
                objTypes.Add("*Line");
                objTypes.Add("*Text");
                //получаем список Id выбранных элементов            
                objectIds = function.getobjectsIds(objTypes);
            }
            //создаем списки объектов
            List<Line> horisontal_lines = new List<Line>();
            List<Line> vertical_lines = new List<Line>();
            List<DBText> text_list = new List<DBText>();
            List<MText> mtext_list = new List<MText>();
            //создаем переменные для определения границ
            double? x_min = null;
            double? y_min = null;
            double? x_max = null;
            double? y_max = null;


            //проходим по объектам и распределяем по спискам
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in objectIds)
                {                
                    Object base_obj = tr.GetObject(id, OpenMode.ForRead) as Object;
                    //получаем объект как линию
                    Line obj = base_obj as Line;
                    if (obj != null)
                    {
                        //получаем проекции начала и конца
                        Point3d point1 = new Point3d(obj.StartPoint.X, obj.StartPoint.Y, 0);
                        Point3d point2 = new Point3d(obj.EndPoint.X, obj.EndPoint.Y, 0);
                        bool using_line = false;
                        //получаем горизонтальную или вертикальную линии (параллельность осям с допуском в 1%)
                        if (Math.Abs(point1.X - point2.X) < (point1.DistanceTo(point2) / 100))
                        {
                            using_line = true;
                            vertical_lines.Add(obj.Clone() as Line);
                        }
                        if (Math.Abs(point1.Y - point2.Y) < (point1.DistanceTo(point2) / 100))
                        {
                            using_line = true;
                            horisontal_lines.Add(obj.Clone() as Line);
                        }
                        if (using_line)
                        {
                            //проверяем начало и конец линии на максимальные/минимальные координаты и устанавливаем их
                            if (x_min == null) x_min = point1.X;
                            else if (point1.X < x_min) x_min = point1.X;
                            if (y_min == null) y_min = point1.Y;
                            else if (point1.Y < y_min) y_min = point1.Y;
                            if (x_max == null) x_max = point1.X;
                            else if (point1.X > x_max) x_max = point1.X;
                            if (y_max == null) y_max = point1.Y;
                            else if (point1.Y > y_max) y_max = point1.Y;

                            if (x_min == null) x_min = point2.X;
                            else if (point2.X < x_min) x_min = point2.X;
                            if (y_min == null) y_min = point2.Y;
                            else if (point2.Y < y_min) y_min = point2.Y;
                            if (x_max == null) x_max = point2.X;
                            else if (point2.X > x_max) x_max = point2.X;
                            if (y_max == null) y_max = point2.Y;
                            else if (point2.Y > y_max) y_max = point2.Y;
                        }                   
                        continue;
                    }
                    //получаем объект как текст
                    DBText obj2 = base_obj as DBText;
                    if (obj2 != null)
                    {
                        text_list.Add(obj2.Clone() as DBText);                    
                        continue;
                    }
                    //получаем объект как полилинию
                    Polyline obj3 = base_obj as Polyline;
                    if (obj3 != null)
                    {
                        //получаем число сегментов полилинии
                        int k = 0;
                        if (obj3.Closed) k = obj3.NumberOfVertices;
                        else k = obj3.NumberOfVertices - 1;
                        //проходим по всем сегментам
                        for (int i = 0; i < k; i++)
                        {
                            //если сегмент линейный
                            if (obj3.GetSegmentType(i) == SegmentType.Line)
                            {
                                //получаем сегмент
                                using (LineSegment3d line3d = obj3.GetLineSegmentAt(i))
                                {
                                    //создаем новую линию из сегмента
                                    Line vv_curve = function.GetLineFromGeLine3d(line3d) as Line;
                                    Point3d point1 = new Point3d(vv_curve.StartPoint.X, vv_curve.StartPoint.Y, 0);
                                    Point3d point2 = new Point3d(vv_curve.EndPoint.X, vv_curve.EndPoint.Y, 0);
                                    //переменная определяющая является ли линия горизонтальной или вертикальной
                                    bool using_line = false;
                                    //получаем горизонтальную или вертикальную линии, устаналиваем статус используемой и добавляем в соответствующий список
                                    if (Math.Abs(point1.X - point2.X) < (point1.DistanceTo(point2) / 100))
                                    {
                                        using_line = true;
                                        vertical_lines.Add(new Line(point1, point2));
                                    }
                                    if (Math.Abs(point1.Y - point2.Y) < (point1.DistanceTo(point2) / 100))
                                    {
                                        using_line = true;
                                        horisontal_lines.Add(new Line(point1, point2));
                                    }
                                    //если линия используемая то проверяем ее на крайние координаты
                                    if (using_line)
                                    {
                                        //проверяем на максимальные/минимальные координаты
                                        if (x_min == null) x_min = point1.X;
                                        else if (point1.X < x_min) x_min = point1.X;
                                        if (y_min == null) y_min = point1.Y;
                                        else if (point1.Y < y_min) y_min = point1.Y;
                                        if (x_max == null) x_max = point1.X;
                                        else if (point1.X > x_max) x_max = point1.X;
                                        if (y_max == null) y_max = point1.Y;
                                        else if (point1.Y > y_max) y_max = point1.Y;

                                        if (x_min == null) x_min = point2.X;
                                        else if (point2.X < x_min) x_min = point2.X;
                                        if (y_min == null) y_min = point2.Y;
                                        else if (point2.Y < y_min) y_min = point2.Y;
                                        if (x_max == null) x_max = point2.X;
                                        else if (point2.X > x_max) x_max = point2.X;
                                        if (y_max == null) y_max = point2.Y;
                                        else if (point2.Y > y_max) y_max = point2.Y;
                                    }
                                }
                            }
                        }                       
                        continue;
                    }
                    //получаем объект как мультитекст
                    MText obj4 = base_obj as MText;
                    if (obj4 != null)
                    {
                        mtext_list.Add(obj4.Clone() as MText);                    
                        continue;
                    }                  
                }
                tr.Commit();
            }
            //если нет минимальных или максимальных координат то значит соответствующий линии отсутствуют и программа прекращает работу
            if (!x_max.HasValue | !x_min.HasValue | !y_max.HasValue | !y_min.HasValue) return;
            //создаем переменные для определения границ
            double x_min_p = x_min.Value;
            double y_min_p = y_min.Value;
            double x_max_p = x_max.Value;
            double y_max_p = y_max.Value;

            //получаем средний габарит таблицы
            double delta_x = (x_max.Value - x_min.Value) / 100;
            double delta_y = (y_max.Value - y_min.Value) / 100;
            //получаем точность определения границ ячеек по горизонтали
            double tab_delta_x = delta_x * Properties.Settings.Default.tab_prec / 100.0;
            //получаем точность определения границ ячеек по вертикали
            double tab_delta_y = delta_y * Properties.Settings.Default.tab_prec / 100.0;
            //получаем точность определения границы соседней ячейки по горизонтали
            double cell_delta = Properties.Settings.Default.cell_prec / 100;
            //создаем списки горизонтальных и вертикальных координат
            //горизонтальные и вертикальные координаты являются координатами углов ячеек в таблице 
            //и получаются из координат начал и концов горизонтальных и верткиальных линии из которых состоит таблица
            List<double> horisontal_coordinates = new List<double>();
            List<double> vertical_coordinates = new List<double>();
            //что бы не задваивать координаты от нескольких линий лежащих на одной прямой или близко к этому \
            //сверяем координаты линий и если они соответствуют уже получанной координате
            //в пределах допуска в 1% от габарита таблицы то эта линия пропускается
            //проходим по горизонтальным линиям и получаем верткиальные координаты границ ячеек
            foreach (Line horisontal_line in horisontal_lines)
            {
                //параметр есть ли в списке координат координата соответствующая этой линии
                bool nenaideno = true;
                //проходим по уже полученным координатам
                foreach (double vertical_coordinate in vertical_coordinates)
                {
                    //проверяем координату Y одной из точек линии
                    //(так как она горизонтальная разницей в координатах начала и конца линии можно пренебречь)
                    if (Math.Abs(horisontal_line.StartPoint.Y - vertical_coordinate) <= tab_delta_y)
                    {
                        //если координата соответствующая текущей в списке уже есть то ставим параметр и прекращаем сравнение
                        nenaideno = false;
                        break;
                    }
                }
                //если соответсвующей координаты не нашлось то добавляем ее в список
                if (nenaideno)
                {
                    vertical_coordinates.Add(horisontal_line.StartPoint.Y);
                }
            }
            //производим аналогичные действия с вертикальными линиями получая список горизонтальных координат границ ячеек
            foreach (Line vertical_line in vertical_lines)
            {
                bool nenaideno = true;
                foreach (double horisontal_coordinate in horisontal_coordinates)
                {
                    if (Math.Abs(vertical_line.StartPoint.X - horisontal_coordinate) <= tab_delta_x)
                    {
                        nenaideno = false;
                        break;
                    }
                }
                if (nenaideno)
                {
                    horisontal_coordinates.Add(vertical_line.StartPoint.X);
                }
            }
            //если вдруг коодинат меньше 2, что является отсутствием ячеек то прекращаем работу
            if (vertical_coordinates.Count < 2 | horisontal_coordinates.Count < 2) return;
            //сортируем полученные координаты
            //что дает возможность проходить по ним как по ячейкам
            vertical_coordinates.Sort();
            horisontal_coordinates.Sort();
            //создаем строки соответсвующие линийм таблицы,
            //столбцы в строках разделены знаком табуляции
            //позже список этих строк будет использоваться для выгрузки в Excel
            //список текстов в ячейке
            List<DBText> cur_texts = new List<DBText>();
            //список мультитекстов в ячейке
            List<MText> cur_mtexts = new List<MText>();
            //список результирующих строк
            List<string> lines = new List<string>();
            //проходим по ячейкам таблицы используя вертикальные и горизонтальные координаты ячеек
            //проходим по строкам таблицы
            //i текущая вертикальная координата которая определяет строку таблицы
            //проходка идет по возрастанию координат, что является проходкой по таблице снизу вверх
            for (int i = 0; i < (vertical_coordinates.Count - 1); i++)
            {
                //создаем переменную для строки с данными в текущей строке таблицы
                string line = string.Empty;
                //счетчик ячеек в объединении
                int cell_int = 0;
                //проходим по столбцам таблицы (j) определяя данные в ячейках соответствующим текущей строке
                for (int j = 0; j < (horisontal_coordinates.Count - 1); j++)
                {
                    //обнуляем списки текстов и мультитекстов в ячейке
                    cur_mtexts.Clear();
                    cur_texts.Clear();
                    //получаем границы ячейки из координат строк и столбцов
                    //координата Х левой стороны ячейки
                    x_min = horisontal_coordinates[j];
                    //координата Х правой стороны ячейки
                    x_max = horisontal_coordinates[j + 1];
                    //координата Y нижней стороны ячейки
                    y_min = vertical_coordinates[i];
                    //координата Y верхней стороны ячейки
                    y_max = vertical_coordinates[i + 1];
                    //ищем реальный конец ячейки так как ячейка может быть объединена по горизонтали
                    //переменная для остановки если найден конец ячейки
                    bool stop = false;
                    //пока не найден конец ячейки если она объединенная
                    while (!stop)
                    {
                        //получаем координаты точки центра правой стороны ячейки
                        Point3d v_point = new Point3d(x_max.Value, (y_min.Value + y_max.Value) / 2, 0);
                        if (cell_delta >= 1) cell_delta = 0.95;
                        double cell_delta_x = (Math.Abs(y_min.Value - y_max.Value)) / 2 * cell_delta;
                        //если вертикальная линия проходит через точку или в пределах допуска таблицы от нее значит это конец ячейки
                        //если такой линии нет значит текущая ячейка и следующая объеденеты в одну
                        //проходим по списку вертикальных линий
                        foreach (Line v_line in vertical_lines)
                        {
                            //создаем копию линии что бы работать в плане
                            using (Line vv_line = (Line)v_line.Clone())
                            {
                                //обнуляем высоты скопированной линии
                                vv_line.StartPoint = new Point3d(vv_line.StartPoint.X, vv_line.StartPoint.Y, 0);
                                vv_line.EndPoint = new Point3d(vv_line.EndPoint.X, vv_line.EndPoint.Y, 0);
                                //получаем точку проекцию на линию
                                Point3d proekc = vv_line.GetClosestPointTo(v_point, false);
                                //получаем расстояние от точки до линии
                                double f_dist = proekc.DistanceTo(v_point);

                                if (f_dist <= cell_delta_x)
                                {
                                    stop = true;
                                    break;

                                }
                            }
                        }
                        //если расстояние в пределах допуска то конец ячейки найден на текущих координатах
                        //останавливаем работу перебора и переходим к дальнейшей работе
                        if (stop) break;
                        //если линия соответствующая концу ячейки не найдена то переходим к следующей ячейке
                        j++;
                        cell_int++;
                        //если дошли до конца таблицы то останавливаем,
                        //значит таблица не закрыта с правой стороны ячейки линией
                        //и текущая координата конца ячейки справа является концом таблицы
                        if (j == (horisontal_coordinates.Count - 1))
                        {
                            break;
                        }
                        //устанвливаем новую координату конца ячейки если справа ячейки еще есть
                        //и переходим к проверке следующей ячейки
                        x_max = horisontal_coordinates[j + 1];
                    }
                    //проходим по всем текстам и мтекстам
                    //получаем тексты и мтексты находящиеся в ячейке
                    foreach (DBText dBText in text_list)
                    {
                        //получаем точку центра области, занятой текстом
                        Point3d point = function.Get_bounds_center(dBText);
                        //если текст в пределах координат текущей ячейки то добавляем этот текст в список текстов ячейки
                        if (point.X < x_max & point.X >= x_min & point.Y < y_max & point.Y >= y_min)
                        {
                            cur_texts.Add(dBText);
                        }
                    }
                    foreach (MText dBText in mtext_list)
                    {
                        Point3d point = function.Get_bounds_center(dBText);
                        if (point.X < x_max & point.X >= x_min & point.Y < y_max & point.Y >= y_min)
                        {
                            cur_mtexts.Add(dBText);
                        }
                    }
                    //создаем переменную, в которой будет записан текст в ячейке
                    string demi_line = string.Empty;
                    //проходим по всем текстам и мтекстам в ячейке
                    //если их несколько то создаем результирующую строку
                    //на основании взаимного расположения текстов 
                    //если текст один то просто записываем его в demi_line
                    if ((cur_texts.Count + cur_mtexts.Count) > 1)
                    {
                        //так как в ячейке последовательность положения текстов 
                        //определена координатами x и y, но может быть небольшой разбег 
                        //в координатах y  текстов расположенных в одной строке 
                        //а так же последовательность по y координате должна идти от большего к меньшему
                        //а по x координате от меньшего к большему
                        //то будем работать с координатами по отдельности

                        //список координат текстов несортированный
                        List<Point3d> position_list_ns = new List<Point3d>();
                        //переменная для суммы высот текстов и средней высоты строки
                        double text_h_s = 0;
                        //проходим по текстам, 
                        foreach (DBText dBText in cur_texts)
                        {
                            //получаем координаты центра текста
                            Point3d point = function.Get_bounds_center(dBText);
                            //записываем координаты в список
                            position_list_ns.Add(point);
                            //добавляем высоту текста к сумме высот текстов
                            text_h_s += dBText.Height;
                        }
                        foreach (MText mText in cur_mtexts)
                        {
                            Point3d point = function.Get_bounds_center(mText);
                            position_list_ns.Add(point);
                            text_h_s += mText.TextHeight;
                        }
                        //получаем среднюю используемую высоту строки как примерно 70% от средней высоты шрифта
                        text_h_s /= ((cur_mtexts.Count + cur_texts.Count) * 1.5);

                        //создаем список в котором будут отсортированные координаты текстов
                        List<Point3d> position_list = new List<Point3d>();
                        //сортируем координаты текстов в ячейке
                        //проходим по списку несортированных координат пока не заполним список отсортированных координат
                        while (position_list_ns.Count != 0)
                        {

                            //переменная позиции самого верхнего текста
                            int pl_i = 0;
                            //проходим по всем текстам и получаем номер самого верхнего
                            for (int pl = 0; pl < position_list_ns.Count; pl++)
                            {
                                if (position_list_ns[pl].Y > position_list_ns[pl_i].Y) pl_i = pl;
                            }
                            //создаем список в котором будут храниться координаты текстов в текущей строке
                            List<Point3d> position_line_list = new List<Point3d>();
                            //добавляем в список самый верхний текст
                            position_line_list.Add(position_list_ns[pl_i]);
                            //проходим по всем текстам
                            for (int pl = 0; pl < position_list_ns.Count; pl++)
                            {
                                //если текст самый верхний пропускаем, он уже есть в списке
                                if (pl == pl_i) continue;
                                //если текст по высоте в пределах параметра text_h_s
                                //то считаем что текст в одной строке с самым верхним
                                //и добавляем в список текстов в одной строке 
                                if (Math.Abs(position_list_ns[pl].Y - position_list_ns[pl_i].Y) < text_h_s)
                                {
                                    position_line_list.Add(position_list_ns[pl]);
                                }
                            }
                            //сортируем тексты по взаимному положению текстов в самой строке                            
                            if (position_line_list.Count > 1)
                            {
                                List<Point3d> position_line_list_v = new List<Point3d>();
                                while (position_line_list.Count > 0)
                                {
                                    int minpos = 0;
                                    for (int t = 1; t < position_line_list.Count; t++)
                                    {
                                        if (position_line_list[t].X < position_line_list[minpos].X) minpos = t;
                                    }
                                    position_line_list_v.Add(position_line_list[minpos]);
                                    position_line_list.RemoveAt(minpos);
                                }
                                position_line_list.AddRange(position_line_list_v);
                            }
                            //добавляем тексты строки в список сортированных координат
                            position_list.AddRange(position_line_list);
                            //удаляем выбранные позиции
                            for (int y = (position_list_ns.Count - 1); y >= 0; y--)
                            {
                                foreach (Point3d y_i in position_list)
                                {
                                    if (y_i.Equals(position_list_ns[y]))
                                    {
                                        position_list_ns.RemoveAt(y);
                                        break;
                                    }
                                }
                            }
                        }
                        //составляем строку в ячейке из последовательности текстов по сортированным координатам
                        //проходим по координатам
                        for (int i_n = 0; i_n < position_list.Count; i_n++)
                        {
                            //устанавливаем разделитель между текстами
                            string razd = " ";
                            //получаем текующя позицию
                            Point3d position = position_list[i_n];
                            //проходим по текстам и мтекстам и ищем текст соответствующий текущей позиции
                            foreach (DBText dBText in cur_texts)
                            {
                                //получаем позицию текста
                                Point3d point = function.Get_bounds_center(dBText);
                                //если текст соответствует текущей позиции то добавляем его в строку ячейки
                                if (point.X.Equals(position.X) & point.Y.Equals(position.Y))
                                {
                                    //если строка пуста то добавляем в нее текст
                                    //если нет то проверяем является ли текст
                                    //текстом этой строки или должен быть перенесен на следующую
                                    if (string.IsNullOrEmpty(demi_line))
                                    {
                                        demi_line = dBText.TextString;
                                    }
                                    else
                                    {
                                        //если позиция текста как минимум вторая
                                        if (i_n > 0)
                                        {
                                            //проверяем находится ли текст на однйо строке с предыдущим,
                                            //если не то разделителем будет переход строки
                                            if (Math.Abs(position_list[i_n].Y - position_list[i_n - 1].Y) > text_h_s)
                                            {
                                                razd = "\r\n";
                                            }
                                        }
                                        //добавляем текст к общей строке ячейки
                                        demi_line = demi_line + razd + dBText.TextString;
                                    }
                                    //останавливаем перебор текстов
                                    break; ;
                                }
                            }
                            //проверяем мтекст аналогично
                            foreach (MText dBText in cur_mtexts)
                            {
                                Point3d point = function.Get_bounds_center(dBText);
                                if (point.X.Equals(position.X) & point.Y.Equals(position.Y))
                                {
                                    if (string.IsNullOrEmpty(demi_line))
                                    {
                                        demi_line = dBText.Text;
                                    }
                                    else
                                    {
                                        if (i_n > 0)
                                        {
                                            if (Math.Abs(position_list[i_n].Y - position_list[i_n - 1].Y) > text_h_s)
                                            {
                                                razd = "\r\n";
                                            }
                                        }
                                        demi_line = demi_line + razd + dBText.Text;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        //просто добавляем текст в переменную строки ячейки если он всего один
                        foreach (DBText dBText in cur_texts)
                        {
                            demi_line = dBText.TextString;
                        }
                        foreach (MText mText in cur_mtexts)
                        {
                            demi_line = mText.Text;
                        }
                    }
                    //обрабалываем слишком длинные тексты без переносов строк
                    //иногда встречаются мтексты без переносов,
                    //где перенос регулируется автокадом через размер окна текста
                    //в этом случае добавляем переносы сами,
                    //так как в ячейку экселя больше 255 символов не влазит и получается очень некрасиво
                    //ограничиваем строку 70 символами                    
                    if (demi_line.Length >= 70)
                    {
                        //добавляем перенос строки если она больше 70 символов
                        demi_line = function.perenos_stroki(demi_line, 70);
                    }
                    //если ячейка первая в списке то добавляем текст ячейки в итоговую строку
                    //если нет то добавляем к итоговой строке через табуляцию, как текста следующей ячейки
                    if ((j - cell_int) == 0)
                    {
                        line = demi_line;
                    }
                    else
                    {
                        line = line + "\t" + demi_line;
                    }
                    while (cell_int > 0)
                    {
                        line += "\t";
                        cell_int--;
                    }
                    //если ячейка последняя то добавляем результирующую строку ряда в список строк
                    if (j >= (horisontal_coordinates.Count - 2))
                    {
                        lines.Add(line);
                    }
                }
            }
            //так как строки у нас созданы по координатам снизу вверх то инвертируем список строк
            lines.Reverse();
            //если вдруг строк не оказалось завершаем работу            
            if (lines.Count == 0) return;


            ///////////////////////////////////работа с экселем///////////////////////
            ////создаем экземпляр прилоежения
            Excel.Application excApp;
            try
            {
                excApp = new Excel.Application();
            }
            catch
            {
                winForm.MessageBox.Show("Excel не найден", "Ошибка", winForm.MessageBoxButtons.OK, winForm.MessageBoxIcon.Exclamation);
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
            //заполняем страницу данными
            for (int i = 0; i < lines.Count; i++)
            {
                string data_line = lines[i];
                //записываем остальные данные
                int j = 1;
                while (data_line.Length > 0)
                {
                    if (data_line.IndexOf("\t") != -1)
                    {
                        string sbstr = data_line.Substring(0, data_line.IndexOf("\t"));
                        data_line = data_line.Remove(0, data_line.IndexOf("\t") + 1);
                        if (!string.IsNullOrEmpty(sbstr))
                        {                            
                            if (sbstr.Contains("\\A1;м{\\H0.7x;\\S3^;}")) sbstr = sbstr.Replace("\\A1;м{\\H0.7x;\\S3^;}", "м3");
                            if (sbstr.Contains("\\A1;м{\\H0.7x;\\S2^;}")) sbstr = sbstr.Replace("\\A1;м{\\H0.7x;\\S2^;}", "м2");
                            if (function.IsNumber(sbstr))
                            {
                                if (sbstr.Contains(",")) sbstr = sbstr.Replace(",", ".");
                                if (sbstr.Contains(",") | sbstr.Contains("."))
                                {
                                    string format = "0,00";
                                    int k = sbstr.Length - sbstr.IndexOf(".") - sbstr.IndexOf(",") - 2;
                                    if (k > 0)
                                    {
                                        format = "0,";
                                        for (int kk = 0; kk < k; kk++)
                                        {
                                            format += "0";
                                        }
                                    }
                                    worksheet.Cells[i + 1, j].NumberFormat = format;
                                }
                            }
                            else
                            {
                                worksheet.Cells[i + 1, j].NumberFormat = "@";
                            }
                            worksheet.Cells[i + 1, j] = sbstr;
                        }
                        j++;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(data_line))
                        {
                            if (data_line.Contains("\\A1;м{\\H0.7x;\\S3^;}")) data_line = data_line.Replace("\\A1;м{\\H0.7x;\\S3^;}", "м3");
                            if (data_line.Contains("\\A1;м{\\H0.7x;\\S2^;}")) data_line = data_line.Replace("\\A1;м{\\H0.7x;\\S2^;}", "м2");
                            if (function.IsNumber(data_line))
                            {
                                if (data_line.Contains(",") | data_line.Contains("."))
                                {
                                    string format = "0,00";
                                    int k = data_line.Length - data_line.IndexOf(".") - data_line.IndexOf(",") - 2;
                                    if (k > 0)
                                    {
                                        format = "0,";
                                        for (int kk = 0; kk < k; kk++)
                                        {
                                            format += "0";
                                        }
                                    }
                                    worksheet.Cells[i + 1, j].NumberFormat = format;
                                }
                                if (data_line.Contains(",")) data_line = data_line.Replace(",", ".");
                            }
                            else
                            {
                                worksheet.Cells[i + 1, j].NumberFormat = "@";
                            }
                            worksheet.Cells[i + 1, j] = data_line;
                            data_line = "";
                        }
                    }
                }
            }
            if (!Properties.Settings.Default.cell_format)
            {
                excApp.UserControl = true;
                excApp.Visible = true;
                return;
            }
            //инвертируем координаты
            vertical_coordinates.Reverse();
            //устанавливаем вертикальные границы          
            //проходим по вертикальным линиям
            foreach (Line vertical_line in vertical_lines)
            {
                //устанавливаем вертикальную переменную ячейки в 0
                int vert_index = 0;
                //проходим по горизонтальным координатам
                for (int i = 0; i < horisontal_coordinates.Count; i++)
                {
                    if (Math.Abs(vertical_line.StartPoint.X - horisontal_coordinates[i]) <= tab_delta_x)
                    {
                        vert_index = i + 1;
                        break;
                    }
                }
                int horisontal_index_v = 0;
                int horisontal_index_n = 0;
                for (int i = 0; i < vertical_coordinates.Count; i++)
                {
                    if (Math.Abs(vertical_line.StartPoint.Y - vertical_coordinates[i]) <= tab_delta_y |
                        Math.Abs(vertical_line.EndPoint.Y - vertical_coordinates[i]) <= tab_delta_y)
                    {
                        if (horisontal_index_v == 0)
                        {
                            horisontal_index_v = i + 1;
                            continue;
                        }
                        else
                        {
                            horisontal_index_n = i;
                            break;
                        }
                    }
                }
                if (vert_index == 0 | horisontal_index_n == 0 | horisontal_index_v == 0) continue;
                Excel.Range range = worksheet.get_Range((Excel.Range)worksheet.Cells[horisontal_index_v, vert_index], (Excel.Range)worksheet.Cells[horisontal_index_n, vert_index]);
                range.Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
            }
            foreach (Line horisontal_line in horisontal_lines)
            {
                int horisontal_index = 0;
                for (int i = 0; i < vertical_coordinates.Count; i++)
                {
                    if (Math.Abs(horisontal_line.StartPoint.Y - vertical_coordinates[i]) <= tab_delta_y)
                    {
                        horisontal_index = i + 1;
                        break;
                    }
                }
                int vertical_index_v = 0;
                int vertical_index_n = 0;
                for (int i = 0; i < horisontal_coordinates.Count; i++)
                {
                    if (Math.Abs(horisontal_line.StartPoint.X - horisontal_coordinates[i]) <= tab_delta_x |
                        Math.Abs(horisontal_line.EndPoint.X - horisontal_coordinates[i]) <= tab_delta_x)
                    {
                        if (vertical_index_v == 0)
                        {
                            vertical_index_v = i + 1;
                            continue;
                        }
                        else
                        {
                            vertical_index_n = i;
                            break;
                        }
                    }
                }
                if (horisontal_index == 0 | vertical_index_n == 0 | vertical_index_v == 0) continue;
                Excel.Range range = worksheet.get_Range((Excel.Range)worksheet.Cells[horisontal_index, vertical_index_n], (Excel.Range)worksheet.Cells[horisontal_index, vertical_index_v]);
                range.Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
            }
            try
            {
                for (int j = 1; j < worksheet.UsedRange.Rows.Count; j++)
                {
                    Excel.Range r_start = null;
                    for (int i = 1; i < worksheet.UsedRange.Columns.Count; i++)
                    {
                        Excel.Range r_end = worksheet.Cells[j, i + 1];
                        if (r_end.Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle == 1 | i == (worksheet.UsedRange.Columns.Count - 1))
                        {
                            if (r_start != null)
                            {
                                r_start = worksheet.get_Range(r_start, (Excel.Range)worksheet.Cells[j, i]);
                                r_start.Merge();
                                r_start = null;
                            }
                        }
                        else
                        {
                            if (r_start == null)
                            {
                                r_start = (Excel.Range)worksheet.Cells[j, i];
                            }
                        }

                    }
                }
            }
            catch
            {
                winForm.MessageBox.Show("ошибка объединения ячеек по горизонтали", "Ошибка", winForm.MessageBoxButtons.OK, winForm.MessageBoxIcon.Exclamation);
            }
            //объединяем ячейки по вертикали
            try
            {
                for (int j = 1; j <= worksheet.UsedRange.Columns.Count; j++)
                {
                    //объявлякем переменную ячейки начала объединенного участка
                    Excel.Range r_start = null;
                    //проходим по колонкам
                    for (int i = 1; i < worksheet.UsedRange.Rows.Count; i++)
                    {
                        //объявляем переменную конца объединенного учатка как следующую ячейку
                        Excel.Range r_end = worksheet.Cells[i + 1, j];
                        //если сверху от нижней ячейки есть граинца значит текущая и предыдущие едины 
                        //переходим к дальнейшей проверке
                        if (r_end.Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle == 1)
                        {
                            //если есть начальная ячейка то интервал состоит как минимум из двух и тогда переходим к объединению
                            if (r_start != null)
                            {
                                //double cell_double;
                                //string cell = string.Empty;
                                Object cell_value = worksheet.Cells[r_start.Row, j].Value;
                                string cell = string.Empty;
                                if (cell_value != null) cell = cell_value.ToString();

                                for (int k = r_start.Row + 1; k <= i; k++)
                                {

                                    //string cell2 = string.Empty;
                                    string cell2 = string.Empty;
                                    cell_value = worksheet.Cells[k, j].Value;
                                    if (cell_value != null) cell2 = cell_value.ToString();
                                    if (!string.IsNullOrEmpty(cell2))
                                    {
                                        if (string.IsNullOrEmpty(cell)) cell = cell2;
                                        else
                                        {
                                            cell = cell + "\r\n" + cell2;
                                        }
                                        worksheet.Cells[k, j].Value = "";
                                    }
                                }
                                worksheet.Cells[r_start.Row, j] = cell;
                                r_start = worksheet.get_Range(r_start, (Excel.Range)worksheet.Cells[i, j]);
                                r_start.Merge();
                                r_start = null;
                            }
                        }
                        else
                        {
                            if (r_start == null)
                            {
                                r_start = (Excel.Range)worksheet.Cells[i, j];
                            }
                        }

                    }
                }
            }
            catch
            {
                winForm.MessageBox.Show("ошибка объединения ячеек по вертикали", "Ошибка", winForm.MessageBoxButtons.OK, winForm.MessageBoxIcon.Exclamation);
            }
            //получаем область таблицы
            Excel.Range used_range = worksheet.UsedRange;
            //получаем пустую ячейку за границами таблицы и ее первоначальные габариты
            Excel.Range blank = worksheet.Cells[(worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count + 2), (worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count + 2)];

            //получаем число строк и столбцов в области таблицы
            int r_c = worksheet.UsedRange.Rows.Count;
            int c_c = worksheet.UsedRange.Columns.Count;

            //устанавливаем ширину ячеек базовую
            used_range.Columns.ColumnWidth = 8.11;
            //устанавливаем высоту столбца базовую
            used_range.Rows.RowHeight = 14.4;
            //устанавливаем в интервале ширину одиночных ячеек по содержимому            
            //проходим по столбцам
            for (int j = 1; j <= c_c; j++)
            {
                //проходим по строкам
                for (int i = 1; i < r_c; i++)
                {
                    //получаем ячейку
                    Excel.Range range = worksheet.Cells[i, j];
                    //проверяем объединенная ячейка или нет
                    if (!range.MergeCells)
                    {
                        //получаем длину ячейки с текстом
                        double max_text_length = this.Max_lenght_in_cell(blank, range);
                        //если длина не больше текущей то переходим к следующей
                        if (max_text_length <= range.ColumnWidth) continue;
                        //устанавливаем длину ячейки по максимальной если она в допустимых пределах
                        if (max_text_length < 250) worksheet.Cells[i, j].ColumnWidth = (max_text_length + 0.2);
                    }
                }
            }
            //проходим по столбцам
            for (int j = 1; j <= c_c; j++)
            {
                //проходим по строкам
                for (int i = 1; i < r_c; i++)
                {
                    //получаем ячейку
                    Excel.Range range = worksheet.Cells[i, j];
                    //проверяем объединенная ячейка или нет
                    if (range.MergeCells)
                    {
                        //получаем объединенный диапазон
                        Excel.Range r_area = (Excel.Range)range.MergeArea;
                        //получаем длину ячейки с текстом
                        double max_text_length = this.Max_lenght_in_cell(blank, range);
                        //устанавливаем длину ячейки по максимальной если она в допустимых пределах
                        //и больше чем сумма длинн ячеек составляющих объединенную
                        if (max_text_length < 250)
                        {
                            //создаем список текущих ширин колонок входящих в объединенный диапазон
                            List<double> col_wid = new List<double>();
                            //создаем переменную суммарной длины
                            double col_widj_sum = 0;
                            //получаем и суммируем длины ячеек
                            for (int v = r_area.Column; v < (r_area.Column + r_area.Columns.Count); v++)
                            {
                                col_widj_sum += worksheet.Cells[1, v].ColumnWidth;
                            }
                            //если суммарная длина меньше чем нужная то распределяем недостачу на все ячейки диапазона
                            if (col_widj_sum < max_text_length)
                            {
                                //получаем часть недостачи для каждой ячейки
                                double inc = (max_text_length - col_widj_sum) / r_area.Columns.Count;
                                //прибавляем недостачу к каждой ячейке и чутка сверху на всякий пожарный
                                for (int v = r_area.Column; v < (r_area.Column + r_area.Columns.Count); v++)
                                {
                                    worksheet.Cells[1, v].ColumnWidth += (inc + 0.02);
                                }
                            }
                        }
                    }
                }
            }
            //проходим по строкам
            for (int i = 1; i <= r_c; i++)
            {
                //проходим по столбцам
                for (int j = 1; j < c_c; j++)
                {
                    //получаем ячейку
                    Excel.Range range = worksheet.Cells[i, j];
                    //проверяем объединенная ячейка или нет
                    if (!range.MergeCells)
                    {
                        //получаем высоту ячейки с текстом
                        double max_text_height = this.Max_height_in_cell(blank, range);
                        //если высота не больше текущей то переходим к следующей
                        if (max_text_height <= range.RowHeight) continue;
                        //устанавливаем высоту ячейки по максимальной если она в допустимых пределах
                        if (max_text_height < 250) worksheet.Cells[range.Row, range.Column].RowHeight = (max_text_height + 0.2);
                    }
                }
            }
            //проходим по строкам
            for (int i = 1; i <= r_c; i++)
            {
                //проходим по столбцам
                for (int j = 1; j < c_c; j++)
                {
                    //получаем ячейку
                    Excel.Range range = worksheet.Cells[i, j];
                    //проверяем объединенная ячейка или нет
                    if (range.MergeCells)
                    {
                        //получаем объединенный диапазон
                        Excel.Range r_area = (Excel.Range)range.MergeArea;
                        //получаем высоту ячейки с текстом
                        double max_text_height = this.Max_height_in_cell(blank, range);
                        //устанавливаем высоту ячейки по максимальной если она в допустимых пределах
                        //и больше чем сумма длинн ячеек составляющих объединенную
                        if (max_text_height < 250)
                        {
                            //создаем список текущих ширин колонок входящих в объединенный диапазон
                            List<double> col_wid = new List<double>();
                            //создаем переменную суммарной длины
                            double col_widj_sum = 0;
                            //получаем и суммируем длины ячеек
                            for (int v = r_area.Row; v < (r_area.Row + r_area.Rows.Count); v++)
                            {
                                col_widj_sum += worksheet.Cells[v, 1].RowHeight;
                            }
                            //если суммарная длина меньше чем нужная то распределяем недостачу на все ячейки диапазона
                            if (col_widj_sum < max_text_height)
                            {
                                //получаем часть недостачи для каждой ячейки
                                double inc = (max_text_height - col_widj_sum) / r_area.Rows.Count;
                                //прибавляем недостачу к каждой ячейке и чутка сверху на всякий пожарный
                                for (int v = r_area.Row; v < (r_area.Row + r_area.Rows.Count); v++)
                                {
                                    worksheet.Cells[v, 1].RowHeight += (inc + 0.02);
                                }
                            }
                        }
                    }
                }
            }
            //возвращаем габариты спец ячейки
            blank.ColumnWidth = 8.11;
            blank.RowHeight = 14.4;
            //выравниваем текст если стоит выравнивание
            if (Properties.Settings.Default.cell_po_centru)
            {
                used_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                used_range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            //если все нормально то передаем управление пользователю
            if (book.Worksheets.Count > 0)
            {
                excApp.UserControl = true;
                excApp.Visible = true;
            }
            else
            {
                //обнуляем книгу и листы
                book = null;
                //закрываем ексель если ничего не нашлось
                excApp.Application.Quit();
                excelProc.Kill();
                excApp = null;
                excelProc = null;
                winForm.MessageBox.Show("Нет данных для выгрузки", "Ошибка", winForm.MessageBoxButtons.OK, winForm.MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// принимает адрес спец ячейки и адрес ячейки с текстом и возвращает максимальную длину текста
        /// </summary>
        /// <param name="blank"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        private double Max_lenght_in_cell(Excel.Range blank, Excel.Range range)
        {
            //получаем содержимое ячейки как текст          
            Object cell_value = range.Value;
            if (cell_value == null) return 0.0;
            string cell = cell_value.ToString();
            //если ячейка пустая то переходим к следующей
            if (string.IsNullOrEmpty(cell)) return 0.0;
            //получаем длину максимальной строки                            
            double max_length = 0;
            if (cell.Contains("\r\n"))
            {
                //получаем подстроку
                string max_text;
                //получаем индексы переноса строк
                List<int> r_n_pos = function.AllIndexesOf(cell, "\r\n");
                //проходим по подстрокам
                for (int v = 0; v <= r_n_pos.Count; v++)
                {
                    if (v == 0)
                    {
                        max_text = cell.Substring(0, r_n_pos[v]);
                    }
                    else
                    {
                        if (v == r_n_pos.Count)
                        {
                            max_text = cell.Substring(r_n_pos[v - 1] + 2);
                        }
                        else
                        {
                            max_text = cell.Substring(r_n_pos[v - 1] + 2, r_n_pos[v] - (r_n_pos[v - 1] + 2));
                        }
                    }
                    //вставляем подстроку в спец ячейку
                    blank.Value = max_text;
                    //устанавливаем ширину ячейки под текст
                    blank.EntireColumn.AutoFit();
                    //если ширина больше максимальной то считаем максимальной ее
                    if (blank.ColumnWidth > max_length) max_length = blank.ColumnWidth;
                }
            }
            else
            {
                //вставляем подстроку в спец ячейку
                blank.Value = cell;
                //устанавливаем ширину ячейки под текст
                blank.EntireColumn.AutoFit();
                //если ширина больше максимальной то считаем максимальной ее
                if (blank.ColumnWidth > max_length) max_length = blank.ColumnWidth;
            }
            blank.Value = "";
            //если ширину ячейки получить не удалось переходим к следующей
            return max_length;

        }
        private double Max_height_in_cell(Excel.Range blank, Excel.Range range)
        {
            //получаем содержимое ячейки как текст          
            Object cell_value = range.Value;
            if (cell_value == null) return 0.0;
            string cell = cell_value.ToString();
            //если ячейка пустая то переходим к следующей
            if (string.IsNullOrEmpty(cell)) return 0;
            //устанавливаем длину строки максимальную что бы текст не переходил на другие строки и не увеличивал высоту
            blank.ColumnWidth = 250;
            //вставляем подстроку в спец ячейку
            blank.Value = cell;
            //устанавливаем высоту ячейки под текст
            blank.EntireRow.AutoFit();
            //получаем высоту строки 
            double max_length = blank.RowHeight;
            blank.Value = "";
            return max_length;
        }

    }
}
