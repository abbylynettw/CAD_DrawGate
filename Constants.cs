// Constants
using System;
using System.IO;
using System.Reflection;
using MultiExcelMultiDoor;

public class Constants
{
    public   string DWG_CORNER_PATH = INIHelper.IniReadValue("FILEPATH", "DWGCornerPath", "无");

    public   string DWG_BLOCK_PATH = INIHelper.IniReadValue("FILEPATH", "DWGBlOCKPATH", "无");

    public   string EXCEL_PATH = INIHelper.IniReadValue("FOlDERPATH", "EXCELCornerPath", "无");
    public   string EXCEL__POLY_PATH = INIHelper.IniReadValue("FOlDERPATH", "EXCELPolyPath", "无");
    public   string DXF_SAVE_PATH = INIHelper.IniReadValue("FOlDERPATH", "DXFPath", "无");
    public   string DWG_SAVE_PATH = INIHelper.IniReadValue("FOlDERPATH", "DWGPath", "无");

    public   string BorderlayerName = INIHelper.IniReadValue("LAYERNAME", "layer", "无");

    public static string assemblyDirectory => Path.GetDirectoryName(Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path));
}