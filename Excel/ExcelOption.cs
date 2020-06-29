// ArxDotNetLesson.Excel.ExcelOption
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

public static class ExcelOption
{
	public static List<string> ChooseExcel()
	{
		List<string> list = new List<string>();
		try
		{
			OpenFileDialog fileDialog = new OpenFileDialog();
			fileDialog.Multiselect = true;
			fileDialog.Title = "请选择表格";
			fileDialog.Filter = "所有文件(*xls*)|*.xls*";
			if (fileDialog.ShowDialog() == DialogResult.OK)
			{
				string[] fileNames = fileDialog.FileNames;
				foreach (string file in fileNames)
				{
					list.Add(file);
				}
				return list;
			}
		}
		catch (Exception e)
		{
			MessageBox.Show(e.ToString());
			return null;
		}
		return list;
	}

	public static string ChooseFolder()
	{
		FolderBrowserDialog dialog = new FolderBrowserDialog();
		dialog.Description = "请选择生成的图纸存放路径";
		string foldPath = "";
		if (dialog.ShowDialog() == DialogResult.OK)
		{
			foldPath = dialog.SelectedPath + "\\";
		}
		return foldPath;
	}

	public static List<GateParameter> ReadExcel(string fileName)
	{
		List<GateParameter> excelist = new List<GateParameter>();
		try
		{
			using (FileStream filesrc = File.OpenRead(fileName))
			{
				if (string.IsNullOrEmpty(fileName))
				{
					return excelist;
				}


				IWorkbook workbook = new XSSFWorkbook((Stream)filesrc);
				for (int i = 0; i < workbook.NumberOfSheets; i++)
				{
					ISheet sheet = workbook.GetSheetAt(i);
					for (int r = 1; r <= sheet.LastRowNum; r++)
					{
                        IRow row = sheet.GetRow(r);
						if (row != null)
						{
                            if (row.GetCell(0) == null || string.IsNullOrWhiteSpace(row.GetCell(0).ToString()))
							{
                                break;
							}
							excelist.Add(new GateParameter
							{
								DwgName = ((row.GetCell(0) == null) ? "" : row.GetCell(0).ToString()),
                                BorderWidth = ((row.GetCell(1) == null) ? 0.0 : StringToDouble(row.GetCell(1).ToString())),
                                BorderHeight = ((row.GetCell(2) == null) ? 0.0 : StringToDouble(row.GetCell(2).ToString())),
                                BorderBaseX = ((row.GetCell(3) == null) ? 0.0 : StringToDouble(row.GetCell(3).ToString())),
                                BorderBaseY = ((row.GetCell(4) == null) ? 0.0 : StringToDouble(row.GetCell(4).ToString())),

                                GateHeight = ((row.GetCell(5) == null) ? 0.0 : StringToDouble(row.GetCell(5).ToString())),
								GateWidth = ((row.GetCell(6) == null) ? 0.0 : StringToDouble(row.GetCell(6).ToString())),
								LCornerBlockName = ((row.GetCell(7) == null) ? "" : row.GetCell(7).ToString()),
								RCornerBlockName = ((row.GetCell(8) == null) ? "" : row.GetCell(8).ToString()),
								LUCornerBlockName = ((row.GetCell(9) == null) ? "" : row.GetCell(9).ToString()),
								RUCornerBlockName = ((row.GetCell(10) == null) ? "" : row.GetCell(10).ToString()),
								FitBlockName = ((row.GetCell(11) == null) ? "" : row.GetCell(11).ToString()),
								FitXoffset = ((row.GetCell(12) == null) ? "" : row.GetCell(12).ToString()),
                                FitYoffse = ((row.GetCell(13) == null) ? "" : row.GetCell(13).ToString()),
                                IsMirrored = ((row.GetCell(14) == null) ? "" : row.GetCell(14).ToString()),

                                RectangleMinX = ((row.GetCell(15) == null) ? 0.0 : StringToDouble(row.GetCell(15).ToString())),
                                RectangleMinY = ((row.GetCell(16) == null) ? 0.0 : StringToDouble(row.GetCell(16).ToString())),
                                RectangleMaxX = ((row.GetCell(17) == null) ? 0.0 : StringToDouble(row.GetCell(17).ToString())),
                                RectangleMaxY = ((row.GetCell(18) == null) ? 0.0 : StringToDouble(row.GetCell(18).ToString())),
                            });
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show("生成错误！" + ex);
		}
		return excelist;
	}
    public static List<GateParameter> ReadPolyDoorExcel(string fileName)
    {
        List<GateParameter> excelist = new List<GateParameter>();
        try
        {
            using (FileStream filesrc = File.OpenRead(fileName))
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    return excelist;
                }


                IWorkbook workbook = new XSSFWorkbook((Stream)filesrc);
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    for (int r = 1; r <= sheet.LastRowNum; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row != null)
                        {
                            if (row.GetCell(0) == null || string.IsNullOrWhiteSpace(row.GetCell(0).ToString()))
                            {
                                break;
                            }
                            excelist.Add(new GateParameter
                            {
                                DwgName = ((row.GetCell(0) == null) ? "" : row.GetCell(0).ToString()),
                                BorderWidth = ((row.GetCell(1) == null) ? 0.0 : StringToDouble(row.GetCell(1).ToString())),
                                BorderHeight = ((row.GetCell(2) == null) ? 0.0 : StringToDouble(row.GetCell(2).ToString())),
                                BorderBaseX = ((row.GetCell(3) == null) ? 0.0 : StringToDouble(row.GetCell(3).ToString())),
                                BorderBaseY = ((row.GetCell(4) == null) ? 0.0 : StringToDouble(row.GetCell(4).ToString())),

                                PolyX = ((row.GetCell(5) == null) ? 0.0 : StringToDouble(row.GetCell(5).ToString())),
                                PolyY = ((row.GetCell(6) == null) ? 0.0 : StringToDouble(row.GetCell(6).ToString())),
                             
                                FitBlockName = ((row.GetCell(7) == null) ? "" : row.GetCell(7).ToString()),
                                FitXoffset = ((row.GetCell(8) == null) ? "" : row.GetCell(8).ToString()),
                                FitYoffse = ((row.GetCell(9) == null) ? "" : row.GetCell(9).ToString()),
                                IsMirrored = ((row.GetCell(10) == null) ? "" : row.GetCell(10).ToString()),

                                RectangleMinX = ((row.GetCell(11) == null) ? 0.0 : StringToDouble(row.GetCell(11).ToString())),
                                RectangleMinY = ((row.GetCell(12) == null) ? 0.0 : StringToDouble(row.GetCell(12).ToString())),
                                RectangleMaxX = ((row.GetCell(13) == null) ? 0.0 : StringToDouble(row.GetCell(13).ToString())),
                                RectangleMaxY = ((row.GetCell(14) == null) ? 0.0 : StringToDouble(row.GetCell(14).ToString())),
                            });
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("生成错误！" + ex);
        }
        return excelist;
    }
    public static double StringToDouble(this string obj)
	{
		double result = 0.0;
		double.TryParse(obj, out result);
		return result;
	}
}
