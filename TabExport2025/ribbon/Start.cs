using Autodesk.AutoCAD.Runtime;

namespace TabExport
{
    public class Start
    {
        [CommandMethod("TableToExcel")]
        public static void TableToExcel()
        {
            TableExportClass.Start();
        }    
    }
}
