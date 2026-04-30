using BaseFunction;
using System;
using System.IO;

namespace TabExport.Settings
{
    public static class Settings
    {
        static Settings()
        {
            Load();
        }

        public static void Load()
        {
            if (BaseXMLClass.GetSerialisationResult(Name, typeof(SettingsClass), false) is SettingsClass model) Default = model;
            else Default = new SettingsClass();            
        }
        public static void Save()
        {
            BaseXMLClass.SetSerialisationResult(Name, Default, false);
        }

        static readonly string Name = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AlzAcadProgramSettings", "TabExport.xml");

        public static SettingsClass Default { get; private set; } = new SettingsClass();
    }

    public class SettingsClass : BaseClass
    {
        public int MaxStringLength { get => _MaxStringLength; set { if (value < 0) return; SetData(ref _MaxStringLength, value); } }
        private int _MaxStringLength = 50;
    }
  
}



