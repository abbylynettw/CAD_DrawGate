// MultiExcelMultiDoor.IOHelper
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

public static class IOHelper
{

    private static List<string> list = new List<string>();

    public static void CreateDirectory(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }
    }

    public static List<string> GetPath(string path)
    {
        DirectoryInfo directoryInfo = new DirectoryInfo(path);
        FileInfo[] files = directoryInfo.GetFiles();
        directoryInfo.GetDirectories();
        list.Clear();
        FileInfo[] array = files;
        FileInfo[] array2 = array;
        foreach (FileInfo fileInfo in array2)
        {
            list.Add(fileInfo.FullName);
        }

        return list;
    }

    /// <summary>
    /// 选择文件
    /// </summary>
    /// <returns></returns>
    public static string SelectPath()
    {
        OpenFileDialog dialog = new OpenFileDialog();
        dialog.Multiselect = true; //该值确定是否可以选择多个文件
        dialog.Title = "请选择文件夹";
        dialog.Filter = "所有文件(*.dwg)|*.dwg";
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            string file = dialog.FileName;
            return file;
        }

        return "";
    }

    public static string GetFile()
    {
        FolderBrowserDialog dialog = new FolderBrowserDialog();
        dialog.Description = "请选择文件路径";
        if (dialog.ShowDialog() == DialogResult.OK)
        {
            string foldPath = dialog.SelectedPath;
            return foldPath + "\\";
        }
        return "";
    }
}