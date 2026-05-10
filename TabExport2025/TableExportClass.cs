using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using BaseFunction;
using System;
using System.Collections.Generic;
using System.Linq;
using TabExport.Data;
using static TabExport.SupportClass;

namespace TabExport
{
    
    public static class TableExportClass
    {     
        public static void Start()
        {
            if (!BaseGetObjectClass.TryGetIntFromUser(out int msl, Settings.Settings.Default.MaxStringLength, 1, null, "Ограничение числа символов в одной строке: ")) return;

            Settings.Settings.Default.MaxStringLength = msl;
            Settings.Settings.Save();

            if (!BaseGetObjectClass.TryGetObjectsIds(out List<ObjectId> ids, new List<Type> { typeof(Line), typeof(Polyline), typeof(DBText), typeof(MText) })) return;
              
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                List<Line> lines = new List<Line>();
                try
                {
                    //получаем исходные данные
                    GetObjects(tr, ids, out List<TextDataClass> texts, out lines);

                    //если нет текстов или линий то прекращаем
                    if (lines.Count == 0 || texts.Count == 0) return;

                    //формируем таблицу
                    TableStructureClass tableStructure = CreateTableStructure(texts, lines);

                    //если таблица не сформировалась то прекращаем
                    if (tableStructure.Columns.Count < 1) return;

                    //создаем эксель документ
                    ExcelClass.CreateExcelDocument.Create(tableStructure);
                }
                finally
                {
                    //очищаем временно созданные объекты
                    foreach (Curve curve in lines) curve?.Dispose();
                }

                tr.Commit();
            }
        }    
        public static void GetObjects(Transaction tr, List<ObjectId> ids, out List<TextDataClass> texts, out List<Line> lines)
        {
            texts = new List<TextDataClass>();
            lines = new List<Line>();

            foreach (ObjectId id in ids)
            {
                Entity entity = tr.GetObject(id, OpenMode.ForRead, false, true) as Entity;

                GetObject(entity, texts, lines);
            }
        }
       
    }
}
