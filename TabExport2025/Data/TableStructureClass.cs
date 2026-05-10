using System.Collections.Generic;

namespace TabExport.Data
{
    public class TableStructureClass
    {       
        public List<RangeClass> Rows { get; set; } = new List<RangeClass>();
        public List<RangeClass> Columns { get; set; } = new List<RangeClass>();
        public DataCellClass[,] Cells { get; set; } = new DataCellClass[1,1];
    }
    public class DataCellClass
    {
        public DataCellClass() { }
        public int Row { get; set; } = 0;
        public int Column { get; set; } = 0;
        public string Value { get; set; } = "";
        public bool VerticalValue { get; set; } = false;
        public int EndRow { get; set; } = 0;
        public int EndColumn { get; set; } = 0;
        public List<TextDataClass> TextDataClasses { get; set; } = new List<TextDataClass> ();
        public bool Checked { get; set; } = false;
        public bool Blocked { get; set; } = false;
    }

    public class RangeClass
    {
        public int Position { get; set; } = 0; 
        public double StartPosition { get; set; } = double.MinValue;
        public double EndPosition { get; set; } = double.MaxValue;
    }

    public class TextDataClass
    {
        public double TextHeight { get; set; } = 0;
        public double X { get; set; } = 0;
        public double Y { get; set; } = 0;
        public string Value { get; set; } = string.Empty;
        public bool VerticalValue { get; set; } = false;
        public Autodesk.AutoCAD.Geometry.PolylineCurve2d Extents { get; set; } 
    }
    
   

}
