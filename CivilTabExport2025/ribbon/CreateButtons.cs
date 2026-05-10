using Autodesk.AutoCAD.Runtime;
using BaseFunction;
using System.Collections.Generic;

namespace CivilTabExport.Ribbon
{
    public class ExampleRibbon : IExtensionApplication
    {
        public void Initialize()
        {
            Buttons();
            CountMenus();
        }
        public void Terminate() { }

        private void Buttons()
        {
            StartEvents startEvents = new StartEvents();
            
            startEvents.Buttons.Add(new Button("nCommand", "Таблица",
                new List<ButtonCommand> { new ButtonCommand("TableToExcel3", "Экспорт таблицы Civil", "Экспортирует в эксель таблицы Civil. Не работает на разделенные таблицы."), }));
           
            startEvents.Initialize();
        }
        private void CountMenus()
        {
        }
    }
}
