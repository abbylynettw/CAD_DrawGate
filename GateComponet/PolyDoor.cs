using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Geometry;

namespace MultiExcelMultiDoor.GateComponet
{
    internal class PolyDoor
    {
        public void Generate(Document doc, List<GateParameter> gatePara, double height, double len, Point3d p)
        {
            try
            {
                using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
                {
                    Polyline polyline = new Polyline();
                    foreach (var gp in gatePara)
                    {
                        polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(gp.PolyX, gp.PolyY), 0.0, 0.0, 0.0);
                    }
                    polyline.Closed = true;
                    polyline.Color = Color.FromRgb(255, 0, 0);
                    var extent = polyline.GeometricExtents;
                    if (height > len)
                    {
                        var pxListLeft = gatePara.Where(g => Math.Abs(g.PolyX - extent.MinPoint.X) < 0.001).OrderBy(g => g.PolyY).ToList();
                        var pxListRight = gatePara.Where(g => Math.Abs(g.PolyX - extent.MaxPoint.X) < 0.001).OrderBy(g => g.PolyY).ToList();
                        var line1 = new Line();
                        line1.StartPoint = new Point3d(pxListLeft.FirstOrDefault().PolyX, pxListLeft.FirstOrDefault().PolyY, 0);
                        line1.EndPoint = new Point3d(line1.StartPoint.X, p.Y, 0);
                        var line2 = new Line();
                        line2.StartPoint=new Point3d(pxListLeft.LastOrDefault().PolyX, pxListLeft.LastOrDefault().PolyY, 0);
                        line2.EndPoint = new Point3d(line2.StartPoint.X, p.Y + height, 0);
                        var line3 = new Line();
                        line3.StartPoint = new Point3d(pxListRight.FirstOrDefault().PolyX, pxListRight.FirstOrDefault().PolyY, 0);
                        line3.EndPoint = new Point3d(line3.StartPoint.X, p.Y, 0);
                        var line4 = new Line();
                        line4.StartPoint = new Point3d(pxListRight.LastOrDefault().PolyX, pxListRight.LastOrDefault().PolyY, 0);
                        line4.EndPoint = new Point3d(line4.StartPoint.X, p.Y + height, 0);
                        line1.ToSpace();
                        line2.ToSpace();
                        line3.ToSpace();
                        line4.ToSpace();
                    }
                    else
                    {
                        var pxListBottom = gatePara.Where(g => Math.Abs(g.PolyY - extent.MinPoint.Y) < 0.001).OrderBy(g => g.PolyX).ToList();
                        var pxListTop = gatePara.Where(g => Math.Abs(g.PolyY - extent.MaxPoint.Y) < 0.001).OrderBy(g => g.PolyX).ToList();
                        var line1 = new Line();
                        line1.StartPoint = new Point3d(pxListBottom.FirstOrDefault().PolyX, pxListBottom.FirstOrDefault().PolyY, 0);
                        line1.EndPoint = new Point3d(p.X, line1.StartPoint.Y, 0);
                        var line2 = new Line();
                        line2.StartPoint = new Point3d(pxListBottom.LastOrDefault().PolyX, pxListBottom.LastOrDefault().PolyY, 0);
                        line2.EndPoint = new Point3d(p.X + len, line2.StartPoint.Y,0);
                        var line3 = new Line();
                        line3.StartPoint = new Point3d(pxListTop.FirstOrDefault().PolyX, pxListTop.FirstOrDefault().PolyY, 0);
                        line3.EndPoint = new Point3d(p.X, line3.StartPoint.Y, 0);
                        var line4 = new Line();
                        line4.StartPoint = new Point3d(pxListTop.LastOrDefault().PolyX, pxListTop.LastOrDefault().PolyY, 0);
                        line4.EndPoint = new Point3d(p.X + len, line4.StartPoint.Y, 0);
                        line1.ToSpace();
                        line2.ToSpace();
                        line3.ToSpace();
                        line4.ToSpace();
                    }
                    polyline.ToSpace(doc.Database);
                    transaction.Commit();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
           
        }
    }
}
