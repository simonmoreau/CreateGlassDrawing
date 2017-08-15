using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;


namespace CreatingDWG
{
    class ACADFunction
    {
        public static void ApplyResultLayer(Document doc, string layerName)
        {
            Database db = doc.Database;
            ObjectIdCollection ids = Selection.SelectEntitiesWithinLayer(doc, layerName);
            string log = layerName;

            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ids)
                {
                    Entity acEnt = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;

                    Dimension dm = acEnt as Dimension;
                    if (dm != null)
                    {
                        if (dm.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 4))
                        {
                            log = log + ";" + "DimensionA";
                            dm.Layer = "01-Cotes_A";
                        }
                        else if (dm.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3))
                        {
                            log = log + ";" + "DimensionB";
                            dm.Layer = "02-Cotes_B";
                        }
                    }
                    else
                    {
                        Line ln = acEnt as Line;
                        if (ln != null)
                        {
                            if (ln.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci,4))
                            {
                                log = log + ";" + "Panel";
                                ln.Layer = "03-Panel_A";
                            }
                            else if (ln.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3))
                            {
                                log = log + ";" + "PanelB";
                                ln.Layer = "04-Panel_B";
                            }
                            else if (ln.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1))
                            {
                                log = log + ";" + "Gothic";
                                ln.Layer = "05-Gothic";
                            }
                            else if (ln.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2))
                            {
                                log = log + ";" + "Cross";
                                ln.Layer = "06-Cross";
                            }
                        }
                        else
                        {
                            Polyline2d pl2d = acEnt as Polyline2d;
                            if (pl2d != null)
                            {
                                if (pl2d.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1))
                                {
                                    log = log + ";" + "Gothic";
                                    pl2d.Layer = "05-Gothic";
                                }
                            }
                            else
                            {
                                Polyline3d pl3d = acEnt as Polyline3d;
                                if (pl3d != null)
                                {
                                    if (pl3d.Color == Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1))
                                    {
                                        log = log + ";" + "Gothic";
                                        pl3d.Layer = "05-Gothic";
                                    }
                                }
                                else
                                {
                                    log = log + ";" + "Annotation";
                                    acEnt.Layer = "07-Annotation";
                                }
                            }

                        }
                    }
                }

                // Save the new object to the database
                acTrans.Commit();
            }

            LogFile.addLine(log);

            //Dispose of the transaction

        }

        public static ObjectIdCollection CopyElements(Document resultDoc, ObjectIdCollection sourceObjectsIdColl, Database sourceDatabase, bool explode) 
        {
            Database acDbNewDoc = resultDoc.Database;
            ObjectIdCollection collection = new ObjectIdCollection();

                // Start a transaction in the new database 
                using (Transaction acTrans = acDbNewDoc.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read           
                    BlockTable acBlkTblNewDoc;
                    acBlkTblNewDoc = acTrans.GetObject(acDbNewDoc.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // Open the Block table record Model space for read
                    BlockTableRecord acBlkTblRecNewDoc;
                    acBlkTblRecNewDoc = acTrans.GetObject(acBlkTblNewDoc[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // Clone the objects to the new database
                    IdMapping acIdMap = new IdMapping();
                    sourceDatabase.WblockCloneObjects(sourceObjectsIdColl, acBlkTblRecNewDoc.ObjectId, acIdMap, DuplicateRecordCloning.Ignore, false);

                    foreach (IdPair idPair in acIdMap)
                    {
                        collection.Add(idPair.Value);
                    }

                    // Save the copied objects to the database
                    acTrans.Commit();
                }
            
            return collection;
        }

        public static void EraseElements(Document doc, ObjectIdCollection elementsToRemoved)
        {
            Database acDbNewDoc = doc.Database;

            // Start a transaction in the new database 
            using (Transaction acTrans = acDbNewDoc.TransactionManager.StartTransaction())
            {

                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acDbNewDoc.BlockTableId,OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (ObjectId id in elementsToRemoved)
                {
                    Entity acObjects;
                    try
                    {
                        acObjects = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;
                        if (acObjects != null)
                        {
                            // Erase objects from the drawing
                            acObjects.Erase();
                        }
                    }
                    catch
                    {
                    }


                }

                // Save the copied objects to the database
                acTrans.Commit();
            }
        }

        public List<string> GetLayersNames(Database acCurDb)
        {
            List<string> layerNames = new List<string>();

            // Start a transaction 
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read 
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                foreach (ObjectId acObjId in acLyrTbl)
                {
                    LayerTableRecord acLyrTblRec;
                    acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForRead) as LayerTableRecord;
                    layerNames.Add(acLyrTblRec.Name);
                }

                // Dispose of the transaction
            }

            return layerNames;
        }

        public static void ChangeName(Document doc, string layerName)
        {
            Database db = doc.Database;
            ObjectIdCollection ids = Selection.SelectNameText(doc, layerName);

            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ids)
                {
                    Entity acEnt = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;

                    MText dm = acEnt as MText;
                    if (dm != null)
                    {
                        dm.Contents = layerName;
                    }
                    else
                    {
                    }
                }

                // Save the new object to the database
                acTrans.Commit();
            }

            //Dispose of the transaction

        }

        public static void ReframeLayouts(Document doc, string layerName)
        {

            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Collect all the paperspace layouts for plotting
                ObjectIdCollection layoutsToPlot = new ObjectIdCollection();

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    if (btr.IsLayout && btr.Name.ToUpper() != BlockTableRecord.ModelSpace.ToUpper())
                    {
                        layoutsToPlot.Add(btrId);
                    }
                }

                foreach (ObjectId btrId in layoutsToPlot)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                    Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

                    ACADFunction.ChangeName(doc, layerName);

                    foreach (ObjectId vpId in lo.GetViewports())
                    {
                        Viewport vp = (Viewport)tr.GetObject(vpId, OpenMode.ForWrite);
                        vp.StandardScale = StandardScaleType.ScaleToFit;
                    }
                }

                tr.Commit();
            }

        }

        public static void PurgeLayers(Document doc)
        {
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction Tx =db.TransactionManager.StartTransaction())
            {
                LayerTable table = Tx.GetObject(db.LayerTableId,OpenMode.ForRead) as LayerTable;

                ObjectIdCollection layIds = new ObjectIdCollection();
                foreach (ObjectId id in table)
                {
                    layIds.Add(id);
                }

                //this function will remove all layers which are used in the drawing file
                db.Purge(layIds);

                foreach (ObjectId id in layIds)
                {
                    DBObject obj = Tx.GetObject(id, OpenMode.ForWrite);
                    try
                    {
                        obj.Erase();
                    }
                    catch
                    {
                        
                    }
                    
                }
                Tx.Commit();
            }
        }
    }
}
