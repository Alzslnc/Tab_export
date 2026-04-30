using Autodesk.AutoCAD.Runtime;
using BaseFunction;
using System.Collections.Generic;

namespace TabExport.Ribbon
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
                new List<ButtonCommand> { new ButtonCommand("TableToExcel", "Экспорт Excel", "Экспортирует в эксель таблицы из выбранных примитивов."), }));  

            startEvents.Initialize();
        }
        private void CountMenus()
        {
        }
    }
}
