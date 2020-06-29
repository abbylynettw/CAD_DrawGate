// Rectangle
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using MultiExcelMultiDoor;
using MultiExcelMultiDoor.GateComponet;

internal class Rectangle : IGateComponet
{
    public void Generate(Document doc, GateParameter gatePara)
    {
        double rectangleMinX = gatePara.RectangleMinX;
        double rectangleMinY = gatePara.RectangleMinY;
        double rectangleMaxX = gatePara.RectangleMaxX;
        double rectangleMaxY = gatePara.RectangleMaxY;
        using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
        {
            Polyline polyline = new Polyline();
            polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(rectangleMinX, rectangleMinY), 0.0, 0.0, 0.0);
            polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(rectangleMaxX, rectangleMinY), 0.0, 0.0, 0.0);
            polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(rectangleMaxX, rectangleMaxY), 0.0, 0.0, 0.0);
            polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(rectangleMinY, rectangleMaxY), 0.0, 0.0, 0.0);
            polyline.Closed = true;
            polyline.Color = Color.FromRgb(0, 255, 0);
            polyline.ToSpace(doc.Database);
            transaction.Commit();
        }
    }
}