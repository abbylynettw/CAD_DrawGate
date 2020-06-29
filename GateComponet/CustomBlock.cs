using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
namespace MultiExcelMultiDoor.GateComponet
{
    public class CustomBlock : IGateComponet
    {
        public void Generate(Document doc, GateParameter gatePara)
        {
            try
            {
                using (var trans = doc.Database.TransactionManager.StartTransaction())
                {
                    if (string.IsNullOrEmpty(gatePara.FitXoffset) || string.IsNullOrEmpty(gatePara.FitYoffse))
                    {
                        return;
                    }
                    var mtx = Matrix3d.Displacement(new Vector3d(gatePara.FitXoffset.StringToDouble(), gatePara.FitYoffse.StringToDouble(), 0));
                    DBHelper.InsertBlockRefByDWGFile(doc, trans, new Constants().DWG_BLOCK_PATH, Point3d.Origin, mtx, gatePara.FitBlockName);
                    trans.Commit();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
          
        }

    }
}
