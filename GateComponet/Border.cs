using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;

namespace MultiExcelMultiDoor.GateComponet
{
    internal class Border : IGateComponet
    {
        public void Generate(Document doc, GateParameter gatePara)
        {
            double borderBaseX = gatePara.BorderBaseX;
            double borderBaseY = gatePara.BorderBaseY;
            double w = gatePara.BorderWidth;
            double h = gatePara.BorderHeight;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                Polyline polyline = new Polyline();
                polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(borderBaseX, borderBaseY), 0.0, 0.0, 0.0);
                polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(borderBaseX+w, borderBaseY), 0.0, 0.0, 0.0);
                polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(borderBaseX+w, borderBaseY+h), 0.0, 0.0, 0.0);
                polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(borderBaseX, borderBaseY+h), 0.0, 0.0, 0.0);
                polyline.Closed = true;
                polyline.LayerId = DBHelper.AddLayer(doc.Database, new Constants().BorderlayerName);
                polyline.ToSpace(doc.Database);
                transaction.Commit();
            }
        }
    }
}
