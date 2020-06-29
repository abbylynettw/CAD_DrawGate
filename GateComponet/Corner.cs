// Corner

using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using MultiExcelMultiDoor;
using MultiExcelMultiDoor.GateComponet;
using System.Windows.Forms;

public class Corner : IGateComponet
{
    private Point3d oPoint = Point3d.Origin;

    public void Generate(Document doc, GateParameter gatePara)
    {
        try
        {
            ObjectId null5 = ObjectId.Null;
            ObjectId null7 = ObjectId.Null;
            ObjectId null9 = ObjectId.Null;
            ObjectId null11 = ObjectId.Null;
            Vector3d vector3 = gatePara.GateHeight * Vector3d.YAxis;
            Vector3d vector2 = gatePara.GateWidth * Vector3d.XAxis;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                null5 = DBHelper.InsertBlockRefByDWGFile(doc, transaction, new Constants().DWG_CORNER_PATH, oPoint, null, gatePara.LCornerBlockName);
                null7 = DBHelper.InsertBlockRefByDWGFile(doc, transaction, new Constants().DWG_CORNER_PATH, oPoint + vector2, null, gatePara.RCornerBlockName);
                null9 = DBHelper.InsertBlockRefByDWGFile(doc, transaction, new Constants().DWG_CORNER_PATH, oPoint + vector3, null, gatePara.LUCornerBlockName);
                null11 = DBHelper.InsertBlockRefByDWGFile(doc, transaction, new Constants().DWG_CORNER_PATH, oPoint + vector3 + vector2, null, gatePara.RUCornerBlockName);
                Entity en7 = transaction.GetObject(null5, OpenMode.ForRead) as Entity;
                Entity en6 = transaction.GetObject(null9, OpenMode.ForRead) as Entity;
                Entity en5 = transaction.GetObject(null7, OpenMode.ForRead) as Entity;
                Entity en4 = transaction.GetObject(null11, OpenMode.ForRead) as Entity;
                Line line7 = new Line(oPoint, oPoint + vector3);
                Line line6 = new Line(oPoint + vector2, oPoint + vector2 + vector3);
                Line line5 = new Line(oPoint, oPoint + vector2);
                Line line4 = new Line(oPoint + vector3, oPoint + vector2 + vector3);
                CutLineStart(en7, ref line7);
                CutLineEnd(en6, ref line7);
                CutLineStart(en5, ref line6);
                CutLineEnd(en4, ref line6);
                CutLineStartX(en7, ref line5);
                CutLineEndX(en5, ref line5);
                CutLineStartX(en6, ref line4);
                CutLineEndX(en4, ref line4);
                line7.Color = Color.FromRgb(255, 0, 0);
                line6.Color = Color.FromRgb(255, 0, 0);
                line5.Color = Color.FromRgb(255, 0, 0);
                line4.Color = Color.FromRgb(255, 0, 0);
                line7.ToSpace(doc.Database);
                line6.ToSpace(doc.Database);
                line5.ToSpace(doc.Database);
                line4.ToSpace(doc.Database);
                transaction.Commit();
            }
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }
       
    }

    private void CutLineStart(Entity en, ref Line line)
    {
        Point3dCollection point3dCollection = new Point3dCollection();
        line.IntersectWith(en, Intersect.OnBothOperands, point3dCollection, 0L, 0L);
        if (point3dCollection.Count == 1)
        {
            line.StartPoint = point3dCollection[0];
        }
        else if (point3dCollection.Count == 2)
        {
            line.StartPoint = ((point3dCollection[0].Y > point3dCollection[1].Y) ? point3dCollection[0] : point3dCollection[1]);
        }
    }

    private void CutLineEnd(Entity en, ref Line line)
    {
        Point3dCollection point3dCollection = new Point3dCollection();
        line.IntersectWith(en, Intersect.OnBothOperands, point3dCollection, 0L, 0L);
        if (point3dCollection.Count == 1)
        {
            line.EndPoint = point3dCollection[0];
        }
        else if (point3dCollection.Count == 2)
        {
            line.EndPoint = ((point3dCollection[0].Y < point3dCollection[1].Y) ? point3dCollection[0] : point3dCollection[1]);
        }
    }

    private void CutLineStartX(Entity en, ref Line line)
    {
        Point3dCollection point3dCollection = new Point3dCollection();
        line.IntersectWith(en, Intersect.OnBothOperands, point3dCollection, 0L, 0L);
        if (point3dCollection.Count == 1)
        {
            line.StartPoint = point3dCollection[0];
        }
        else if (point3dCollection.Count == 2)
        {
            line.StartPoint = ((point3dCollection[0].X > point3dCollection[1].X) ? point3dCollection[0] : point3dCollection[1]);
        }
    }

    private void CutLineEndX(Entity en, ref Line line)
    {
        Point3dCollection point3dCollection = new Point3dCollection();
        line.IntersectWith(en, Intersect.OnBothOperands, point3dCollection, 0L, 0L);
        if (point3dCollection.Count == 1)
        {
            line.EndPoint = point3dCollection[0];
        }
        else if (point3dCollection.Count == 2)
        {
            line.EndPoint = ((point3dCollection[0].X < point3dCollection[1].X) ? point3dCollection[0] : point3dCollection[1]);
        }
    }
}
