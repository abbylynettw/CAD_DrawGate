// Command
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using MultiExcelMultiDoor.GateComponet;
using MultiExcelMultiDoor;

public class Command : IExtensionApplication
{
    private string exportPath = "";

    [CommandMethod("SZ")]
    public void Settings()
    {
       Application.ShowModelessDialog(new SettingForm());
    }

    [CommandMethod("PLSCM", CommandFlags.Session)]
    public void GenerateDoor()
    {
        Editor editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        try
        {
            List<string> path = IOHelper.GetPath(new Constants().EXCEL_PATH);
            if (path != null && path.Count > 0)
            {
                List<GateParameter> list = new List<GateParameter>();
                foreach (string item5 in path)
                {
                    List<GateParameter> collection = ExcelOption.ReadExcel(item5);
                    list.AddRange(collection);
                }
                exportPath = new Constants().DXF_SAVE_PATH + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "\\";
                IOHelper.CreateDirectory(exportPath);
                foreach (IGrouping<string, GateParameter> gr in from e in list group e by e.DwgName)
                {
                    string fileName = exportPath + gr.Key + ".dxf";
                    string a = "否";
                    double factor = 0.0;
                    Document document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Add(gr.Key);
                    using (document.LockDocument())
                    {
                        bool flag = false;
                        foreach (GateParameter gp in gr)
                        {
                            if (!((string.IsNullOrEmpty(gp.LCornerBlockName) || string.IsNullOrEmpty(gp.LCornerBlockName) || string.IsNullOrEmpty(gp.LCornerBlockName) || string.IsNullOrEmpty(gp.LCornerBlockName)) | flag))
                            {
                                new Border().Generate(document, gp);
                                new Corner().Generate(document, gp);
                                new Rectangle().Generate(document, gp);
                                flag = true;
                                a = gp.IsMirrored;
                                factor = gp.GateWidth;
                            }
                            new CustomBlock().Generate(document, gp);
                        }
                        if (a == "是")
                        {
                            Line3d line = new Line3d(Point3d.Origin + Vector3d.XAxis * factor, Point3d.Origin + Vector3d.XAxis * factor + Vector3d.YAxis);
                            DBHelper.MatrixEntitiesInModelSpace(Matrix3d.Displacement(-Vector3d.XAxis * factor) * Matrix3d.Mirroring(line));
                        }
                        document.Database.DxfOut(fileName, 16, DwgVersion.AC1021);
                    }
                }
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.SendCommand("CloseAllDwg\n");
            }
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage("生成失败！" + ex?.ToString());
        }
    }

    [CommandMethod("SCDWG", CommandFlags.Session)]
    public void GenerateDwgDoor()
    {
        Editor editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        try
        {
            List<string> path = IOHelper.GetPath(new Constants().EXCEL_PATH);
            if (path != null && path.Count > 0)
            {
                List<GateParameter> list = new List<GateParameter>();
                foreach (string item5 in path)
                {
                    List<GateParameter> collection = ExcelOption.ReadExcel(item5);
                    list.AddRange(collection);
                }
                exportPath = new Constants().DWG_SAVE_PATH + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "\\";
                IOHelper.CreateDirectory(exportPath);
                foreach (IGrouping<string, GateParameter> item4 in from e in list
                                                                   group e by e.DwgName)
                {
                    string fileName = exportPath + item4.Key + ".dwg";
                    string a = "否";
                    double factor = 0.0;
                    Document document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Add(item4.Key);
                    using (document.LockDocument())
                    {
                        bool flag = false;
                        foreach (GateParameter item3 in item4)
                        {
                            if (!((string.IsNullOrEmpty(item3.LCornerBlockName) || string.IsNullOrEmpty(item3.LCornerBlockName) || string.IsNullOrEmpty(item3.LCornerBlockName) || string.IsNullOrEmpty(item3.LCornerBlockName)) | flag))
                            {
                                new Corner().Generate(document, item3);
                                new Rectangle().Generate(document, item3);
                                flag = true;
                                a = item3.IsMirrored;
                                factor = item3.GateWidth;
                            }
                            new CustomBlock().Generate(document, item3);
                        }
                        if (a == "是")
                        {
                            Line3d line = new Line3d(Point3d.Origin + Vector3d.XAxis * factor, Point3d.Origin + Vector3d.XAxis * factor + Vector3d.YAxis);
                            DBHelper.MatrixEntitiesInModelSpace(Matrix3d.Displacement(-Vector3d.XAxis * factor) * Matrix3d.Mirroring(line));
                        }
                        document.Database.SaveAs(fileName, DwgVersion.AC1021);
                    }
                }
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.SendCommand("CloseAllDwg\n");
            }
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage("生成失败！" + ex?.ToString());
        }
    }

    [CommandMethod("SCDDX", CommandFlags.Session)]
    public void GeneratePlineDoor()
    {
        Editor editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        try
        {
            List<string> path = IOHelper.GetPath(new Constants().EXCEL__POLY_PATH);
            if (path != null && path.Count > 0)
            {
                List<GateParameter> list = new List<GateParameter>();
                foreach (string item5 in path)
                {
                    List<GateParameter> collection = ExcelOption.ReadPolyDoorExcel(item5);
                    list.AddRange(collection);
                }
                exportPath = new Constants().DXF_SAVE_PATH + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "\\";
                IOHelper.CreateDirectory(exportPath);
                foreach (IGrouping<string, GateParameter> gr in from e in list group e by e.DwgName)
                {
                    string fileName = exportPath + gr.Key + ".dxf";
                    string a = "否";
                    double factor = 0.0;
                    Document document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Add(gr.Key);
                    using (document.LockDocument())
                    {
                        bool flag = false;
                        List<GateParameter> gpList = new List<GateParameter>();
                        double borderHeight = 0;
                        double borderWidth = 0;
                        Point3d pBorder=new Point3d();
                        foreach (GateParameter gp in gr)
                        {
                            if (!flag)
                            {
                                new Border().Generate(document, gp);
                                new Rectangle().Generate(document, gp);
                                flag = true;
                                a = gp.IsMirrored;
                                factor = gp.GateWidth;
                                borderHeight = gp.BorderHeight;
                                borderWidth = gp.BorderWidth;
                                pBorder = new Point3d(gp.BorderBaseX, gp.BorderBaseY, 0);
                            }
                            new CustomBlock().Generate(document, gp);
                            gpList.Add(gp);
                        }
                        new PolyDoor().Generate(document, gpList, borderHeight, borderWidth, pBorder);
                        if (a == "是")
                        {
                            Line3d line = new Line3d(Point3d.Origin + Vector3d.XAxis * factor, Point3d.Origin + Vector3d.XAxis * factor + Vector3d.YAxis);
                            DBHelper.MatrixEntitiesInModelSpace(Matrix3d.Displacement(-Vector3d.XAxis * factor) * Matrix3d.Mirroring(line));
                        }
                        document.Database.DxfOut(fileName, 16, DwgVersion.AC1021);
                    }
                }
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.SendCommand("CloseAllDwg\n");
            }
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage("生成失败！" + ex?.ToString());
        }
    }

    [CommandMethod("SCDDXDWG", CommandFlags.Session)]
    public void GeneratePlineDoorDWG()
    {
        Editor editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
        try
        {
            List<string> path = IOHelper.GetPath(new Constants().EXCEL__POLY_PATH);
            if (path != null && path.Count > 0)
            {
                List<GateParameter> list = new List<GateParameter>();
                foreach (string item5 in path)
                {
                    List<GateParameter> collection = ExcelOption.ReadPolyDoorExcel(item5);
                    list.AddRange(collection);
                }
                exportPath = new Constants().DWG_SAVE_PATH + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "\\";
                IOHelper.CreateDirectory(exportPath);
                foreach (IGrouping<string, GateParameter> gr in from e in list group e by e.DwgName)
                {
                    string fileName = exportPath + gr.Key + ".dwg";
                    string a = "否";
                    double factor = 0.0;
                    Document document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Add(gr.Key);
                    using (document.LockDocument())
                    {
                        bool flag = false;
                        List<GateParameter> gpList = new List<GateParameter>();
                        double borderHeight = 0;
                        double borderWidth = 0;
                        Point3d pBorder = new Point3d();
                        foreach (GateParameter gp in gr)
                        {
                            if (!flag)
                            {
                                new Border().Generate(document, gp);
                                new Rectangle().Generate(document, gp);
                                flag = true;
                                a = gp.IsMirrored;
                                factor = gp.GateWidth;
                                borderHeight = gp.BorderHeight;
                                borderWidth = gp.BorderWidth;
                                pBorder = new Point3d(gp.BorderBaseX, gp.BorderBaseY, 0);
                            }
                            new CustomBlock().Generate(document, gp);
                            gpList.Add(gp);
                        }
                        new PolyDoor().Generate(document, gpList, borderHeight, borderWidth, pBorder);
                        if (a == "是")
                        {
                            Line3d line = new Line3d(Point3d.Origin + Vector3d.XAxis * factor, Point3d.Origin + Vector3d.XAxis * factor + Vector3d.YAxis);
                            DBHelper.MatrixEntitiesInModelSpace(Matrix3d.Displacement(-Vector3d.XAxis * factor) * Matrix3d.Mirroring(line));
                        }
                        document.Database.SaveAs(fileName, DwgVersion.AC1021);
                    }
                }
                Application.DocumentManager.MdiActiveDocument.SendCommand("CloseAllDwg\n");
            }
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage("生成失败！" + ex?.ToString());
        }
    }
    [CommandMethod("CloseAllDwg", CommandFlags.Session)]
    public void CloseDwgAll()
    {
        DocumentCollection documentManager = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager;
        int num = 0;
        foreach (Document item in documentManager)
        {
            num++;
            if (num != 1)
            {
                if (item.CommandInProgress != "" && item.CommandInProgress != "CloseAllDwg")
                {
                    item.SendCommand("\u0003\u0003");
                }
                if (documentManager.MdiActiveDocument != item)
                {
                    documentManager.MdiActiveDocument = item;
                }
                item.CloseAndDiscard();
            }
        }
     Application.ShowAlertDialog("全部生成成功！");
    }


    [CommandMethod("AddDefaultContextMenu")]
    public void AddCustomMenu()
    {
        ContextMenuExtension obj = new ContextMenuExtension
        {
            Title = "自定义菜单"
        };
        MenuItem menuItem = new MenuItem("批量生成Dxf文件");
        menuItem.Click += mi_Click;
        MenuItem menuItemDwg = new MenuItem("批量生成Dwg文件");
        menuItemDwg.Click += miDwg_Click;
        MenuItem menuItemPline = new MenuItem("批量生成多段线门Dxf");
        menuItemPline.Click += menuItemPline_Click;
        MenuItem menuItemPlineDWG= new MenuItem("批量生成多段线门DWG");
        menuItemPlineDWG.Click += menuItemPlineDWG_Click;
        MenuItem settingItem = new MenuItem("设置");
        settingItem.Click += settingItem_Click;
        obj.MenuItems.Add(menuItem);
        obj.MenuItems.Add(menuItemDwg);
        obj.MenuItems.Add(menuItemPline);
        obj.MenuItems.Add(menuItemPlineDWG);
        obj.MenuItems.Add(settingItem);
        Autodesk.AutoCAD.ApplicationServices.Application.AddDefaultContextMenuExtension(obj);
    }

    private void mi_Click(object sender, EventArgs e)
    {
        Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        if ((sender as MenuItem).Text == "批量生成Dxf文件")
        {
            mdiActiveDocument.SendStringToExecute("PLSCM\n", true, false, true);
        }
    }
    private void miDwg_Click(object sender, EventArgs e)
    {
        Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        if ((sender as MenuItem).Text == "批量生成Dwg文件")
        {
            mdiActiveDocument.SendStringToExecute("SCDWG\n", true, false, true);
        }
    }
    private void menuItemPline_Click(object sender, EventArgs e)
    {
        Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        if ((sender as MenuItem).Text == "批量生成多段线门Dxf")
        {
            mdiActiveDocument.SendStringToExecute("SCDDX\n", true, false, true);
        }
    }
    private void menuItemPlineDWG_Click(object sender, EventArgs e)
    {
        Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        if ((sender as MenuItem).Text == "批量生成多段线门DWG")
        {
            mdiActiveDocument.SendStringToExecute("SCDDXDWG\n", true, false, true);
        }
    }
    private void settingItem_Click(object sender, EventArgs e)
    {
        Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        if ((sender as MenuItem).Text == "设置")
        {
            mdiActiveDocument.SendStringToExecute("SZ\n", true, false, true);
        }
    }
    public void RemoveMenu()
    {
    }

    public void Initialize()
    {
        AddCustomMenu();
        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("自定义菜单 已加载。");
    }

    public void Terminate()
    {
    }
}
