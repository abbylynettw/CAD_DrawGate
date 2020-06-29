// MultiExcelMultiDoor.DBHelper
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

public static class DBHelper
{
    public static ObjectId ToSpace(this Entity ent, Database db = null, string space = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectId id = ObjectId.Null;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            id = mdlSpc.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            trans.Commit();
        }
        return id;
    }

    public static void SendCommand(this Document doc, params string[] args)
    {
        Type AcadDocument = Type.GetTypeFromHandle(Type.GetTypeHandle(doc.GetAcadDocument()));
        try
        {
            AcadDocument.InvokeMember("SendCommand", BindingFlags.InvokeMethod, null, doc.GetAcadDocument(), args);
        }
        catch
        {
        }
    }

    public static Point2d toPoint2d(this Point3d p)
    {
        return new Point2d(p.X, p.Y);
    }

    public static ObjectIdCollection ToSpace(this IEnumerable<Entity> ents, Database db = null, string space = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectIdCollection ids = new ObjectIdCollection();
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            foreach (Entity ent in ents)
            {
                ids.Add(mdlSpc.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            trans.Commit();
        }
        return ids;
    }

    public static void ForEach(this IEnumerable<Entity> ents, Action<Entity> act)
    {
        foreach (Entity ent in ents)
        {
            act(ent);
        }
    }

    public static void ForEach(Action<Entity> act, Database db = null, string space = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
            foreach (ObjectId id in mdlSpc)
            {
                Entity ent = trans.GetObject(id, OpenMode.ForWrite) as Entity;
                act(ent);
            }
            trans.Commit();
        }
    }

    public static ObjectId ToBlockDefinition(this IEnumerable<Entity> ents, string blockName, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectId id = ObjectId.Null;
        BlockTableRecord blkDef = new BlockTableRecord();
        blkDef.Name = blockName;
        foreach (Entity ent in ents)
        {
            blkDef.AppendEntity(ent);
        }
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
            id = blkTbl.Add(blkDef);
            trans.AddNewlyCreatedDBObject(blkDef, true);
            trans.Commit();
        }
        return id;
    }

    public static ObjectId Insert(ObjectId blkDefId, Matrix3d transform, Database db = null, string space = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectId id = ObjectId.Null;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            BlockReference blkRef = new BlockReference(Point3d.Origin, blkDefId);
            blkRef.BlockTransform = transform;
            id = mdlSpc.AppendEntity(blkRef);
            trans.AddNewlyCreatedDBObject(blkRef, true);
            BlockTableRecord blkDef = trans.GetObject(blkDefId, OpenMode.ForRead) as BlockTableRecord;
            if (blkDef.HasAttributeDefinitions)
            {
                foreach (ObjectId subId in blkDef)
                {
                    if (subId.ObjectClass.Equals(RXObject.GetClass(typeof(AttributeDefinition))))
                    {
                        AttributeDefinition attrDef = trans.GetObject(subId, OpenMode.ForRead) as AttributeDefinition;
                        AttributeReference attrRef = new AttributeReference();
                        attrRef.SetAttributeFromBlock(attrDef, transform);
                        blkRef.AttributeCollection.AppendAttribute(attrRef);
                    }
                }
            }
            trans.Commit();
        }
        return id;
    }

    public static ObjectId Insert(string name, Matrix3d transform, Database db = null, string space = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectId id = ObjectId.Null;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (!blkTbl.Has(name))
            {
                return ObjectId.Null;
            }
            id = blkTbl[name];
        }
        return Insert(id, transform, db, space);
    }

    public static ObjectId GetSymbol(ObjectId tblId, string name, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            SymbolTable tbl = trans.GetObject(tblId, OpenMode.ForRead) as SymbolTable;
            if (tbl.Has(name))
            {
                return tbl[name];
            }
        }
        return ObjectId.Null;
    }

    public static ObjectId ToTable(this SymbolTableRecord record, ObjectId tblId, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectId id = ObjectId.Null;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            SymbolTable tbl = trans.GetObject(tblId, OpenMode.ForWrite) as SymbolTable;
            if (tbl.Has(record.Name))
            {
                return tbl[record.Name];
            }
            tbl.Add(record);
            trans.AddNewlyCreatedDBObject(record, true);
            trans.Commit();
        }
        return id;
    }

    public static bool ModifySymbol(ObjectId tblId, string name, Action<SymbolTableRecord> act, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            SymbolTable tbl = trans.GetObject(tblId, OpenMode.ForRead) as SymbolTable;
            if (!tbl.Has(name))
            {
                return false;
            }
            SymbolTableRecord symbol = trans.GetObject(tbl[name], OpenMode.ForWrite) as SymbolTableRecord;
            act(symbol);
            trans.Commit();
        }
        return true;
    }

    public static bool RemoveSymbol(ObjectId tblId, string name, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            SymbolTable tbl = trans.GetObject(tblId, OpenMode.ForWrite) as SymbolTable;
            if (!tbl.Has(name))
            {
                return false;
            }
            SymbolTableRecord inUse2 = null;
            ObjectId inUseId = ObjectId.Null;
            if (tbl is LayerTable)
            {
                inUseId = db.Clayer;
            }
            else if (tbl is TextStyleTable)
            {
                inUseId = db.Textstyle;
            }
            else if (tbl is DimStyleTable)
            {
                inUseId = db.Dimstyle;
            }
            else if (tbl is LinetypeTable)
            {
                inUseId = db.Celtype;
            }
            if (inUseId.IsValid)
            {
                inUse2 = (trans.GetObject(inUseId, OpenMode.ForRead) as SymbolTableRecord);
                if (inUse2.Name.ToUpper() == name.ToUpper())
                {
                    return false;
                }
            }
            DBObject record = trans.GetObject(tbl[name], OpenMode.ForWrite);
            if (record.IsErased)
            {
                return false;
            }
            ObjectIdCollection idCol = new ObjectIdCollection
            {
                record.ObjectId
            };
            db.Purge(idCol);
            if (idCol.Count == 0)
            {
                return false;
            }
            record.Erase();
            trans.Commit();
        }
        return true;
    }

    public static void AttachXData(this DBObject obj, string app, IEnumerable<TypedValue> datas, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        if (GetSymbol(db.RegAppTableId, app, db) == ObjectId.Null)
        {
            ToTable(new RegAppTableRecord
            {
                Name = app
            }, db.RegAppTableId, db);
        }
        ResultBuffer rb = new ResultBuffer();
        rb.Add(new TypedValue(1001, app));
        foreach (TypedValue data in datas)
        {
            rb.Add(data);
        }
        obj.XData = rb;
    }

    public static void AttachXData(this ObjectId objId, string app, IEnumerable<TypedValue> datas, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject obj = trans.GetObject(objId, OpenMode.ForWrite);
            AttachXData(obj, app, datas, db);
            trans.Commit();
        }
    }

    public static TypedValue[] GetXData(this DBObject obj, string app)
    {
        return obj.GetXDataForApplication(app)?.AsArray()?.Skip(1)?.ToArray();
    }

    public static TypedValue[] GetXData(this ObjectId objId, string app, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject obj = trans.GetObject(objId, OpenMode.ForRead);
            return GetXData(obj, app);
        }
    }

    public static TypedValue? SetXData(this DBObject obj, string app, int idx, TypedValue newVal)
    {
        TypedValue[] valArr = obj.GetXDataForApplication(app)?.AsArray();
        if (valArr != null && idx + 1 < valArr.Length)
        {
            TypedValue oldVal = valArr[idx + 1];
            valArr[idx + 1] = newVal;
            obj.XData = new ResultBuffer(valArr);
            return oldVal;
        }
        return null;
    }

    public static TypedValue? SetXData(this ObjectId objId, string app, int idx, TypedValue newVal, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject obj = trans.GetObject(objId, OpenMode.ForRead);
            return SetXData(obj, app, idx, newVal);
        }
    }

    public static void SetExtData(this DBObject obj, string key, DBObject data, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        if (!obj.ExtensionDictionary.IsValid)
        {
            obj.CreateExtensionDictionary();
        }
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBDictionary dict = trans.GetObject(obj.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
            dict.SetAt(key, data);
        }
    }

    public static void SetExtData(this ObjectId objId, string key, DBObject data, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject obj = trans.GetObject(objId, OpenMode.ForWrite);
            SetExtData(obj, key, data, db);
        }
    }

    public static ObjectId GetExtData(this DBObject obj, string key, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        if (obj.ExtensionDictionary.IsValid)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                DBDictionary dict = trans.GetObject(obj.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                if (dict.Contains(key))
                {
                    return dict.GetAt(key);
                }
            }
        }
        return ObjectId.Null;
    }

    public static ObjectId GetExtData(this ObjectId objId, string key, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject obj = trans.GetObject(objId, OpenMode.ForRead);
            return GetExtData(obj, key, db);
        }
    }

    public static void ModifyExtData(this DBObject obj, Action<DBObject> act, string key, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        ObjectId dataId = GetExtData(obj, key, db);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject data = trans.GetObject(dataId, OpenMode.ForWrite);
            act(data);
        }
    }

    public static void ModifyExtData(this ObjectId objId, Action<DBObject> act, string key, Database db = null)
    {
        db = (db ?? Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database);
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            DBObject obj = trans.GetObject(objId, OpenMode.ForRead);
            ModifyExtData(obj, act, key, db);
        }
    }

    public static BlockTableRecord GetBlock(Document doc, Transaction transaction, string blockname)
    {
        BlockTable bt = (BlockTable)transaction.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        foreach (ObjectId item in bt)
        {
            BlockTableRecord btr = item.GetObject(OpenMode.ForRead) as BlockTableRecord;
            if (btr != null && btr.Name == blockname)
            {
                return btr;
            }
        }
        return null;
    }

    public static bool DelBlock(Document doc, Transaction transaction, BlockTableRecord BlockTableRecord_)
    {
        try
        {
            if (BlockTableRecord_.ObjectId != ObjectId.Null)
            {
                DBObject obj = transaction.GetObject(BlockTableRecord_.ObjectId, OpenMode.ForWrite);
                if (obj != null)
                {
                    obj.Erase();
                    return true;
                }
                return false;
            }
            return false;
        }
        catch (Autodesk.AutoCAD.Runtime.Exception)
        {
            return false;
        }
    }

    public static BlockTableRecord GetBlock(Document doc, Transaction transaction, string blockName, string dwgPath)
    {
        BlockTableRecord block = GetBlock(doc, transaction, blockName);
        if (block == null)
        {
            try
            {
                string dwgBlockName = string.Empty;
                using (Database newDB = new Database(false, true))
                {
                    newDB.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, false, "");
                    newDB.CloseInput(true);
                    bool isequal = false;
                    using (Transaction tr = newDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(newDB.BlockTableId, OpenMode.ForRead);
                        foreach (ObjectId item in bt)
                        {
                            BlockTableRecord btr2 = item.GetObject(OpenMode.ForRead) as BlockTableRecord;
                            if (btr2 != null && btr2.Name.Equals(blockName))
                            {
                                isequal = true;
                            }
                        }
                        foreach (ObjectId item2 in bt)
                        {
                            BlockTableRecord btr = item2.GetObject(OpenMode.ForRead) as BlockTableRecord;
                            if (btr != null && !isequal && btr.Name.IndexOf("TD_") == 0)
                            {
                                dwgBlockName = btr.Name;
                                break;
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(dwgBlockName))
                    {
                        BlockTableRecord nowblock3 = GetBlock(doc, transaction, dwgBlockName);
                        if (nowblock3 == null)
                        {
                            ObjectId idBTR2 = doc.Database.Insert(blockName, newDB, false);
                            BlockTableRecord newbtr2 = transaction.GetObject(idBTR2, OpenMode.ForRead) as BlockTableRecord;
                            nowblock3 = GetBlock(doc, transaction, dwgBlockName);
                            DelBlock(doc, transaction, newbtr2);
                        }
                        block = nowblock3;
                    }
                    else if (isequal)
                    {
                        BlockTableRecord nowblock2 = GetBlock(doc, transaction, blockName);
                        if (nowblock2 == null)
                        {
                            ObjectId idBTR = doc.Database.Insert(blockName + Guid.NewGuid().ToString("N"), newDB, false);
                            BlockTableRecord newbtr = transaction.GetObject(idBTR, OpenMode.ForRead) as BlockTableRecord;
                            nowblock2 = GetBlock(doc, transaction, blockName);
                            DelBlock(doc, transaction, newbtr);
                            block = nowblock2;
                        }
                    }
                    else
                    {
                        block = (doc.Database.Insert(blockName, newDB, false).GetObject(OpenMode.ForRead) as BlockTableRecord);
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                throw ex;
            }
        }
        return block;
    }

    public static ObjectId InsertBlockRefByDWGFile(Document doc, Transaction transaction, string dwgPath, Point3d point, Matrix3d? mtx, string blockName = "")
    {
        ObjectId objectId2 = ObjectId.Null;
        blockName = (string.IsNullOrEmpty(blockName) ? Guid.NewGuid().ToString("N") : Regex.Replace(blockName, "[\\u005c\\u003a]", "X"));
        BlockTableRecord block = GetBlock(doc, transaction, blockName, dwgPath);
        BlockTable bt = (BlockTable)transaction.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        using (BlockReference bref = new BlockReference(point, block.ObjectId))
        {
            if (mtx.HasValue)
            {
                bref.TransformBy(mtx.Value);
            }
            objectId2 = btr.AppendEntity(bref);
            transaction.AddNewlyCreatedDBObject(bref, true);
            return objectId2;
        }
    }

    public static void MatrixEntitiesInModelSpace(Matrix3d matrix)
    {
        using (Transaction transaction = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction())
        {
            BlockTable blockTable = (BlockTable)transaction.GetObject(HostApplicationServices.WorkingDatabase.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId item in (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead))
            {
                Entity obj = transaction.GetObject(item, OpenMode.ForWrite) as Entity;
                obj.TransformBy(matrix);
                obj.DowngradeOpen();
            }
            transaction.Commit();
        }
    }

    public static ObjectId AddLayer(this Database db,string layerName)
    {
        //打开层表
        LayerTable lt = (LayerTable) db.LayerTableId.GetObject(OpenMode.ForRead);
        //如果不存在名为layerName的图层，则新建一个图层
        if (!lt.Has((layerName)))
        {
            LayerTableRecord ltr=new LayerTableRecord();//定义一个新的层表记录
            ltr.Name = layerName;//设置图层名
            lt.UpgradeOpen();//切换层表的状态为写以添加新的图层
            lt.Add(ltr);
            db.TransactionManager.AddNewlyCreatedDBObject(ltr, true);
            lt.DowngradeOpen();
        }
        return lt[layerName];
    }
}
