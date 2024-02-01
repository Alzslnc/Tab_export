using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;

namespace test2
{
    /// <summary>
    /// командный класс для тестовой командыtest10
    /// </summary>
    public class commandclass
    {
        /// <summary>
        /// командный метод для запуска команды
        /// </summary>
        [CommandMethod("test10")]
        public void Runcommand2()
        {
            //ссылка на активный документ
            Document adoc = Application.DocumentManager.MdiActiveDocument;
            //если документ не получен выходит из метода
            if (adoc == null)
                return;
            //создание ссылки на базу данных документа(чертежа)
            Database db = adoc.Database;
            //получение ID таблицы слоев из документа
            ObjectId layertabid = db.LayerTableId;
            ////создание списка (пустого) для сбора имен слоев чертежа
            List<string> LayersNames = new List<string>();
            //создание списка для сбора даных о слоях
           
              

            //запуск транзакции
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //получение объекта таблицы слоев через ID таблицы
                LayerTable layerTable = tr.GetObject(layertabid, OpenMode.ForRead) as LayerTable;
                //для каждого объекта в таблице слоев 
                foreach (ObjectId layerTableRecordId in layerTable)
                {
                    //получаем объект - запись о слое в таблице 
                    LayerTableRecord layerTableRecord = tr.GetObject(layerTableRecordId, OpenMode.ForRead) as LayerTableRecord;
                    //получаем имя слоя и добавляем  в список
                    LayersNames.Add(layerTableRecord.Name);
                }
                //подтверждение транзакции
                tr.Commit();
            }
            //получение редактора документа
            Editor ed = adoc.Editor;
            //для каждой строки в списке названий слоев чертежа
            foreach (string layerName in LayersNames)
            {
                //выводим название в командную строку автокада
                ed.WriteMessage("\n" + layerName);
            }

            List
            ed.WriteMessage("\n Всего слоев - " + LayersNames.Count);
        }
        /// <summary>
        /// создаем структуру
        /// </summary>
        struct LayerData
        {
            //имя слоя
            internal string LayerName;
            //включчен слой или нет
            internal bool LayerIsOn;
            //заморожен слой или нет
            internal bool layerIsFrozen;
        }
    }
}
