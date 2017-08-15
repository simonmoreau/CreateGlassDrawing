using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using System.IO;
using System.Windows.Forms;

namespace CreatingDWG
{
    class MainFunctions
    {
        //private bool _print = true;
        private static string _templateFolder;

        public static void LoopOnVentelles()
        {
            OpenFileDialog filesDialog = new OpenFileDialog();
            DialogResult result = filesDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string folderName = Path.GetDirectoryName(filesDialog.FileNames[0]);
                _templateFolder = folderName + @"\Files\";

                DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                foreach (string fileName in filesDialog.FileNames)
                {
                    Document sourceDoc = acDocMgr.Open(fileName, false);
                    //Active the document
                    if (acDocMgr.MdiActiveDocument != sourceDoc)
                    {
                        acDocMgr.MdiActiveDocument = sourceDoc;
                    }

                    OpenDocument(sourceDoc, acDocMgr, folderName);

                    sourceDoc.CloseAndDiscard();
                }
            }
        }

        public static void OpenDocument(Document sourceDoc, DocumentCollection acDocMgr, string folderName)
        {
            Database sourceDB = sourceDoc.Database;

            //Create Export Directory
            string exportName = Path.GetFileNameWithoutExtension(sourceDoc.Name);
            string exportPath = folderName + @"\" + exportName;
            System.IO.DirectoryInfo di = Directory.CreateDirectory(exportPath);
            System.IO.DirectoryInfo diDWG = Directory.CreateDirectory(di.FullName + @"\DWG");
            System.IO.DirectoryInfo diDXF = Directory.CreateDirectory(di.FullName + @"\DXF");
            System.IO.DirectoryInfo diADXF = Directory.CreateDirectory(di.FullName + @"\DXF FACE A");
            System.IO.DirectoryInfo diBDXF = Directory.CreateDirectory(di.FullName + @"\DXF FACE B");
            System.IO.DirectoryInfo diPDF = Directory.CreateDirectory(di.FullName + @"\PDF");

            int i = 0;

            using (Transaction acTrans = sourceDB.TransactionManager.StartTransaction())
            {
                // This example returns the layer table for the current database
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(sourceDB.LayerTableId, OpenMode.ForRead) as LayerTable;

                string dwgFlatTemplate = _templateFolder + "template_Plan.dwt";
                string dwgBombedTemplate = _templateFolder + "template_Bombed.dwt";
                string dxfTemplate = _templateFolder + "acad.dwt";

                Document newDoc;
                Document bombedDoc = acDocMgr.Add(dwgBombedTemplate);
                Document planDoc = acDocMgr.Add(dwgFlatTemplate);
                Document dxfDoc = acDocMgr.Add(dxfTemplate);

                // Step through the Layer table
                foreach (ObjectId acObjId in acLyrTbl)
                {
                    LayerTableRecord acLyrTblRec;
                    acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForRead) as LayerTableRecord;

                    //Active the source document
                    if (acDocMgr.MdiActiveDocument != sourceDoc)
                    {
                        acDocMgr.MdiActiveDocument = sourceDoc;
                    }

                    //Retrive Elements by Layers
                    string panelName = acLyrTblRec.Name;
                    ObjectIdCollection collection = Selection.SelectEntities(sourceDoc, panelName);
                    ObjectIdCollection checkingCollection = Selection.CheckingCollection(sourceDoc, panelName);

                    if (collection.Count != 0)
                    {
                        if (checkingCollection.Count == 0)
                        {
                            newDoc = planDoc;
                        }
                        else
                        {
                            newDoc = bombedDoc;
                        }

                        //Lock the new doc
                        using (DocumentLock Aclock = newDoc.LockDocument())
                        {
                            //Paste the object in it
                            ObjectIdCollection pastedElementsIds = ACADFunction.CopyElements(newDoc, collection, sourceDB, false);

                            //Active the document
                            if (acDocMgr.MdiActiveDocument != newDoc)
                            {
                                acDocMgr.MdiActiveDocument = newDoc;
                            }

                            ACADFunction.ApplyResultLayer(newDoc, panelName);

                            ACADFunction.ReframeLayouts(newDoc, panelName);


                            ////Print it - Deprecated
                            //while (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                            //{
                            //    System.Threading.Thread.Sleep(500);
                            //}

                            //if (i < 5)
                            //{

                            //}
                            //Tools.MultiSheetPlot(newDoc, diPDF, panelName);

                            //Purge unused layer
                            ACADFunction.PurgeLayers(newDoc);

                            //Save It
                            Tools.SaveDoc(newDoc, diDWG, panelName);

                            //Remove object for the next run
                            ACADFunction.EraseElements(newDoc, pastedElementsIds);
                        }

                        //Active the source document
                        if (acDocMgr.MdiActiveDocument != sourceDoc)
                        {
                            acDocMgr.MdiActiveDocument = sourceDoc;
                        }

                        //Select the face A polyline
                        ObjectIdCollection cyanPolyline = Selection.SelectCyanLines(sourceDoc, panelName);

                        //Seclect the face B polyline
                        ObjectIdCollection greenPolyline = Selection.SelectGreenLines(sourceDoc, panelName);
                       
                        //Active the dxf document
                        if (acDocMgr.MdiActiveDocument != dxfDoc)
                        {
                            acDocMgr.MdiActiveDocument = dxfDoc;
                        }

                        //Lock the new doc
                        using (DocumentLock Aclock = dxfDoc.LockDocument())
                        {
                            //Paste the object in it
                            ObjectIdCollection pastedpolylinesIds = ACADFunction.CopyElements(dxfDoc, cyanPolyline, sourceDB, true);

                            //Purge unused layer
                            ACADFunction.PurgeLayers(dxfDoc);

                            //Save it
                            if (checkingCollection.Count != 0)
                            {
                                Tools.SaveDoc(dxfDoc, diADXF, panelName+"A");
                            }
                            else
                            {
                                Tools.SaveDoc(dxfDoc, diDXF, panelName);
                            }

                            //remove the elements for the next run
                            ACADFunction.EraseElements(dxfDoc, pastedpolylinesIds);
                        }

                        //Check for panel curvature
                        if (checkingCollection.Count != 0)
                        {
                            //Lock the new doc
                            using (DocumentLock Aclock = dxfDoc.LockDocument())
                            {
                                //Paste the object in it
                                ObjectIdCollection pastedpolylinesIds = ACADFunction.CopyElements(dxfDoc, greenPolyline, sourceDB, true);

                                //Purge unused layer
                                ACADFunction.PurgeLayers(dxfDoc);

                                //Save it
                                Tools.SaveDoc(dxfDoc, diBDXF, panelName+"B");

                                //remove the elements for the next run
                                ACADFunction.EraseElements(dxfDoc, pastedpolylinesIds);
                            }
                        }
                    }

                    i++;

                    string logPath = _templateFolder + @"\" + exportName + @".txt";
                    File.WriteAllLines(logPath, LogFile.LogList.ToArray());
                }
            }
        }
    }
}
