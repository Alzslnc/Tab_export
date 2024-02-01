using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Tab_export
{
    internal class Function
    {
        /// <summary>
        /// получение линии из сегмента полилинии
        /// </summary>
        /// <param name="line3d"></param>
        /// <returns></returns>
        public Line GetLineFromGeLine3d(LineSegment3d line3d)
        {
            Line line = new Line(line3d.StartPoint, line3d.EndPoint);                      
            return line;
        }     
        /// <summary>
        /// получение Id объектов множественного выбора
        /// </summary>
        /// <param name="objectTypes">список типов объектов  для выбора в виде DFX названий</param>
        /// <returns></returns>
        public List<ObjectId> getobjectsIds(List<string> objectTypes)
        {
            //создаем строку с типами объектов для фильтра
            string objectTypesAll = "";
            foreach (string objectType in objectTypes)
            {
                objectTypesAll = objectTypesAll + objectType + ",";
            }
            objectTypesAll = objectTypesAll.Substring(0, (objectTypesAll.Length - 1));
            //создаем фильтр
            TypedValue[] values = new TypedValue[]
                {
                new TypedValue((int)DxfCode.Start,objectTypesAll)
                };
            SelectionFilter filter = new SelectionFilter(values);

            //создаем списко Id
            List<ObjectId> objectIds = new List<ObjectId>();
            //запускаем множественный выбор рамкой             
            PromptSelectionResult pResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection(filter);
            //если происходит отмена возвращаем пустой список
            if (pResult.Status == PromptStatus.Cancel)
            {
                objectIds.Clear();
                return objectIds;
            }
            //записываем Id выбранных объектов в список и возвращаем его
            if (pResult.Status == PromptStatus.OK) objectIds.AddRange(pResult.Value.GetObjectIds());
            return objectIds;
        }
        /// <summary>
        /// проверяет строку на численность
        /// </summary>
        /// <param name="value">строка</param>
        /// <returns></returns>
        public bool IsNumber(String value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;
            else
                return double.TryParse(value.Trim(), System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out _);
        }
        public List<int> AllIndexesOf(string str, string value)
        {
            
            List<int> indexes = new List<int>();
            if (String.IsNullOrEmpty(value)) return indexes;
            for (int index = 0; ; index += value.Length)
            {
                index = str.IndexOf(value, index);
                if (index == -1)
                    return indexes;
                indexes.Add(index);
            }
        }
        public Point3d Get_bounds_center(Entity entity)
        {
            Extents3d extents3D;
            if (entity is DBText) extents3D = entity.GeometricExtents;
            else extents3D = MtextRealExtents(entity as MText).Value;                    
            return new Point3d((extents3D.MinPoint.X + extents3D.MaxPoint.X) / 2, (extents3D.MinPoint.Y + extents3D.MaxPoint.Y) / 2, 0);
        }
        /// <summary>
        /// возвращаем реальный габарит MText
        /// </summary>
        /// <param name="mTextId"></param>
        /// <returns></returns>
        private Extents3d? MtextRealExtents(MText mText)
        {
            if (mText != null)
            {
                Point3d point = mText.Location;
                Plane plane = new Plane(point, mText.Normal);
                Vector3d vx = plane.Normal.GetPerpendicularVector().GetNormal().TransformBy(Matrix3d.Rotation(mText.Rotation, plane.Normal, point)).GetNormal(); ;
                Vector3d vy = vx.TransformBy(Matrix3d.Rotation(Math.PI / 2, plane.Normal, point)).GetNormal();
                double h = mText.ActualHeight;
                double w = mText.ActualWidth;
                //получаем нижний левый угол текста
                switch (mText.Attachment)
                {
                    case AttachmentPoint.TopLeft:
                        point = point - vy * h;
                        break;
                    case AttachmentPoint.MiddleCenter:
                        point = point - vy * h / 2 - vx * w / 2;
                        break;
                    case AttachmentPoint.TopCenter:
                        point = point - vy * h - vx * w / 2;
                        break;
                    case AttachmentPoint.TopRight:
                        point = point - vy * h - vx * w;
                        break;
                    case AttachmentPoint.MiddleLeft:
                        point = point - vy * h / 2;
                        break;
                    case AttachmentPoint.MiddleRight:
                        point = point - vy * h / 2 - vx * w;
                        break;
                    case AttachmentPoint.BottomLeft:
                        break;
                    case AttachmentPoint.BottomCenter:
                        point = point - vx * w / 2;
                        break;
                    case AttachmentPoint.BottomRight:
                        point = point - vx * w;
                        break;
                }
                //получаем точки 4 углов в wcs
                //достаточно перспективные данные, дают 4 реальных угла текста а не прямоугольник области
                List<Point3d> points = new List<Point3d>
                {
                    point,
                    point + vx * w + vy * h,
                    point + vx * w,
                    point + vy * h,
                };
                //получаем все координаты точек
                List<double> x = new List<double>();
                List<double> y = new List<double>();
                foreach (Point3d p in points)
                {
                    x.Add(p.X);
                    y.Add(p.Y);
                }
                //возвращаем новые габариты
                return new Extents3d(new Point3d(x.Min(), y.Min(), 0).Project(plane, Vector3d.ZAxis), new Point3d(x.Max(), y.Max(), 0).Project(plane, Vector3d.ZAxis));
            }
            return null;
        }
        public string perenos_stroki(string str, int str_l)
        {
            if (str.Length > str_l)
            {
                List<string> lines = new List<string>();
                while (str.Length > 0)
                {
                    if (str.Contains("\r\n"))
                    {
                        lines.Add(str.Substring(0, str.IndexOf("\r\n")));
                        str = str.Substring(str.IndexOf("\r\n") + 2);
                    }
                    else
                    {
                        lines.Add(str);
                        str = string.Empty;
                    }
                }
                foreach (string line in lines)
                {
                    if (line.Length > str_l)
                    {
                        List<string> words = new List<string>();
                        words.AddRange(line.Split(new Char[] { ' ' }));
                        string n_line = string.Empty;
                        int i = 0;                        
                        foreach (string word in words)
                        {
                            i += word.Length;
                            if (i > str_l)
                            {
                                i = 0;
                                n_line += "\r\n" + word;
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(n_line)) n_line = word;
                                else n_line += " " + word;
                            }
                        }
                        if (string.IsNullOrEmpty(str))
                        {
                            str = n_line;
                        }
                        else
                        {
                            str += "\r\n" + n_line;
                        }
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(str))
                        {
                            str = line;
                        }
                        else
                        {
                            str += "\r\n" + line;
                        }
                    }
                }

            }
            return str;

        }      

    }
}