using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace MultiExcelMultiDoor.GateComponet
{
    /// <summary>
    /// 门的组件接口
    /// </summary>
    public interface IGateComponet
    {
        void Generate(Document doc, GateParameter gatePara);
    }
}
