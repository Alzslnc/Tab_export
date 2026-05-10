using Autodesk.AutoCAD.DatabaseServices;
using BaseFunction;
using System.Collections.Generic;
using TabExport.Data;
using static TabExport.SupportClass;


namespace CivilTabExport
{
    internal class CivilTableExportClass
    {
        public static void Start()
        {
            if (!BaseGetObjectClass.TryGetIntFromUser(out int msl, TabExport.Settings.Settings.Default.MaxStringLength, 1, null, "Ограничение числа символов в одной строке: ")) return;

            TabExport.Settings.Settings.Default.MaxStringLength = msl;
            TabExport.Settings.Settings.Save();            

            if (!BaseGetObjectClass.TryGetobjectId(out ObjectId id, typeof(Autodesk.Civil.DatabaseServices.Table), "Выберите таблицу Civil", true)) return;

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBObjectCollection collection = new DBObjectCollection();
                List<Line> lines = new List<Line>();
                Autodesk.Civil.DatabaseServices.Table table = null;
                try
                {
                    //получаем выбранную таблицу
                    table = tr.GetObject(id, OpenMode.ForRead, false, true).Clone() as Autodesk.Civil.DatabaseServices.Table;
                    if (table != null)
                    {
                        table.Explode(collection);

                        collection = GetExplodedObjects(collection);

                        GetObjects(collection, out List<TextDataClass> texts, out lines);

                        //если нет текстов или линий то прекращаем
                        if (lines.Count == 0 || texts.Count == 0) return;

                        //формируем таблицу
                        TableStructureClass tableStructure = CreateTableStructure(texts, lines);

                        //если таблица не сформировалась то прекращаем
                        if (tableStructure.Columns.Count < 1) return;

                        //создаем эксель документ
                        TabExport.ExcelClass.CreateExcelDocument.Create(tableStructure);
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

        private static DBObjectCollection GetExplodedObjects(DBObjectCollection dBObjectCollection)
        {
            List<BlockReference> refs = new List<BlockReference>();
            for (int i = dBObjectCollection.Count - 1; i >= 0; i--)
            {
                if (dBObjectCollection[i] is BlockReference reference)
                {
                    refs.Add(reference);
                    dBObjectCollection.RemoveAt(i);
                }
            }

            if (refs.Count > 0) 
            {
                DBObjectCollection newCollection = new DBObjectCollection();
                foreach (BlockReference blockReference in refs)
                {
                    blockReference.Explode(newCollection);
                    blockReference?.Dispose();
                }
                foreach (DBObject newObject in GetExplodedObjects(newCollection)) dBObjectCollection.Add(newObject);
            }

            return dBObjectCollection;
        }
    }
}
