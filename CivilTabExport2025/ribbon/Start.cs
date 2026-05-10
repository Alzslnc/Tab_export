using Autodesk.AutoCAD.Runtime;

namespace CivilTabExport
{
    public class Start
    {
        [CommandMethod("TableToExcel3")]
        public static void TableToExcel3()
        {
            CivilTableExportClass.Start();
        }
       
    }
}
