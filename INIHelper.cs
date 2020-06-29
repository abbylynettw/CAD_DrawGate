using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MultiExcelMultiDoor
{
    public static class INIHelper
    {
        static string strIniFilePath="";  // ini配置文件路径  

         static INIHelper()
        {
            strIniFilePath = Constants.assemblyDirectory + @"\Setting.ini";
        }      
        // 返回0表示失败，非0为成功  
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        // 返回取得字符串缓冲区的长度  
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern long GetPrivateProfileString(string section, string key, string strDefault, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern bool CloseHandle(IntPtr hObject);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public static readonly IntPtr HFILE_ERROR = new IntPtr(-1);
     

        /// <summary>
        /// 验证文件占用状态
        /// by wjs 20170523
        /// </summary>
        /// <param name="filename">完整文件名</param>
        /// <returns>true：文件被占用，false：文件未被占用</returns>
        public static bool validateFileState(string filename)
        {
            if (File.Exists(filename))
            {
                IntPtr vHandle = _lopen(filename, OF_READWRITE | OF_SHARE_DENY_NONE);
                if (vHandle == HFILE_ERROR)
                {
                    return true;
                }
                CloseHandle(vHandle);
            }
            return false;
        }
        public static string GetMD5HashFromFile(string fileName)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                FileStream file = new FileStream(fileName, FileMode.Open);
                System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                byte[] retVal = md5.ComputeHash(file);
                file.Close();
                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                }
            }
            catch (Exception ex)
            {                
            }
            return sb.ToString();
        }

        /// <summary>
        /// 读取INI文件
        /// </summary>
        /// <param name="Section"></param>
        /// <param name="Key"></param>
        /// <returns></returns>
        public  static string IniReadValue(string Section, string Key, string DefaultValue)
        {
            StringBuilder temp = new StringBuilder(255);
            long i = GetPrivateProfileString(Section, Key, DefaultValue, temp, 255, strIniFilePath);
            return temp.ToString();
        }


        /// <summary>
        /// 写INI文件
        /// </summary>
        /// <param name="Section"></param>
        /// <param name="Key"></param>
        /// <param name="Value"></param>
        public static void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, strIniFilePath);
        }
    }    
}
