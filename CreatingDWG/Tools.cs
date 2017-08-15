using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System.IO;

namespace CreatingDWG
{
    class Tools
    {
        static public void MultiSheetPlot(Document doc, DirectoryInfo diPDF, string panelName)
        {
            Editor ed = doc.Editor;
            Database db = doc.Database;

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                PlotInfo pi = new PlotInfo();
                PlotInfoValidator piv = new PlotInfoValidator();
                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;

                // A PlotEngine does the actual plotting (can also create one for Preview)

                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                {
                    PlotEngine pe = PlotFactory.CreatePublishEngine();
                    using (pe)
                    {
                        // Collect all the paperspace layouts for plotting
                        ObjectIdCollection layoutsToPlot = new ObjectIdCollection();
                        List<BlockTableRecord> btrListToBeSorted = new List<BlockTableRecord>();

                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                            if (btr.IsLayout && btr.Name.ToUpper() != BlockTableRecord.ModelSpace.ToUpper())
                            {
                                layoutsToPlot.Add(btrId);
                            }
                        }

                        //Sort the layout by name
                        btrListToBeSorted.Sort((x, y) => string.Compare(x.Name, y.Name));

                        foreach (BlockTableRecord btr in btrListToBeSorted)
                        {
                            layoutsToPlot.Add(btr.Id);
                        }

                        // Create a Progress Dialog to provide info and allow thej user to cancel
                        PlotProgressDialog ppd =new PlotProgressDialog(false,layoutsToPlot.Count,true);

                        using (ppd)
                        {
                            int numSheet = 1;

                            foreach (ObjectId btrId in layoutsToPlot)
                            {
                                BlockTableRecord btr =(BlockTableRecord)tr.GetObject(btrId,OpenMode.ForWrite);
                                Layout lo =(Layout)tr.GetObject(btr.LayoutId,OpenMode.ForRead);

                                // We need a PlotSettings object based on the layout settings which we then customize

                                PlotSettings ps =new PlotSettings(lo.ModelType);
                                ps.CopyFrom(lo);

                                // The PlotSettingsValidator helps create a valid PlotSettings object
                                PlotSettingsValidator psv =PlotSettingsValidator.Current;

                                // We'll plot the extents, centered and scaled to fit
                                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                                psv.SetUseStandardScale(ps, true);
                                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                                psv.SetPlotCentered(ps, true);

                                
                                // We'll use the standard PDF, as this supports multiple sheets
                                //Think about changing the DWG To PDF settings to custom value for setting up printing margin
                                //psv.SetPlotConfigurationName(ps,"DWFx ePlot (XPS Compatible).pc3", "ANSI_A_(8.50_x_11.00_Inches)");
                                //psv.SetPlotConfigurationName(ps, "DWG To PDF_Custom.pc3", "ISO_A4_(210.00_x_297.00_MM)");
                                psv.SetPlotConfigurationName(ps, "DWG To PDF_Custom.pc3", "ISO_expand_A4_(210.00_x_297.00_MM)");

                                // We need a PlotInfo object linked to the layout

                                pi.Layout = btr.LayoutId;

                                // Make the layout we're plotting current
                                LayoutManager.Current.CurrentLayout =lo.LayoutName;

                                //Change the name in it
                                ACADFunction.ChangeName(doc, panelName);

                                    foreach (ObjectId vpId in lo.GetViewports())
                                    {
                                        Viewport vp = (Viewport)tr.GetObject(vpId, OpenMode.ForWrite);
                                        vp.StandardScale = StandardScaleType.ScaleToFit;
                                        //vp.CustomScale = 10;
                                    }

                                // We need to link the PlotInfo to the PlotSettings and then validate it

                                pi.OverrideSettings = ps;
                                piv.Validate(pi);

                                if (numSheet == 1)
                                {
                                    ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
                                    ppd.set_PlotMsgString( PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                                    ppd.set_PlotMsgString( PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                                    ppd.set_PlotMsgString( PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                                    ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption,"Sheet Progress");
                                    ppd.LowerPlotProgressRange = 0;
                                    ppd.UpperPlotProgressRange = 100;
                                    ppd.PlotProgressPos = 0;

                                    // Let's start the plot, at last

                                    ppd.OnBeginPlot();
                                    ppd.IsVisible = true;
                                    pe.BeginPlot(ppd, null);

                                    // We'll be plotting a single document
                                    // Let's plot to file
                                    pe.BeginDocument(pi, doc.Name, null, 1, true, diPDF.FullName +@"\"+ panelName + ".pdf");
                                }

                                // Which may contain multiple sheets

                                ppd.StatusMsgString = "Plotting " + doc.Name.Substring(doc.Name.LastIndexOf("\\") + 1) +" - sheet " + numSheet.ToString() + " of " + layoutsToPlot.Count.ToString();

                                ppd.OnBeginSheet();

                                ppd.LowerSheetProgressRange = 0;
                                ppd.UpperSheetProgressRange = 100;
                                ppd.SheetProgressPos = 0;

                                PlotPageInfo ppi = new PlotPageInfo();
                                pe.BeginPage(ppi,pi,(numSheet == layoutsToPlot.Count),null);
                                pe.BeginGenerateGraphics(null);
                                ppd.SheetProgressPos = 50;
                                pe.EndGenerateGraphics(null);

                                // Finish the sheet
                                pe.EndPage(null);
                                ppd.SheetProgressPos = 100;
                                ppd.OnEndSheet();
                                numSheet++;
                            }

                            // Finish the document

                            pe.EndDocument(null);

                            // And finish the plot

                            ppd.PlotProgressPos = 100;
                            ppd.OnEndPlot();
                            pe.EndPlot(null);
                        }
                    }
                }
                else
                {
                    ed.WriteMessage(
                      "\nAnother plot is in progress."
                    );
                }

                tr.Commit();
            }
        }

        public static void SaveDoc(Document doc, DirectoryInfo di, string panelName )
        {
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string extension;
            if (di.Name == "DWG")
            {
                extension = ".dwg";
            }
            else
            {
                extension = ".dxf";
            }

            string dwgFile = di.FullName + "\\" + CheckFileName(panelName, di) + extension;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //Set the doc current
                if (di.Name == "DWG")
                {
                    db.SaveAs(dwgFile, true, DwgVersion.AC1021, db.SecurityParameters);
                }
                else
                {
                    db.DxfOut(dwgFile, 16, DwgVersion.AC1021);
                }
            }
        }

        public static string CheckFileName(string shortName, DirectoryInfo di)
        {
            //check for an already existing version
            string[] filesNames = Directory.GetFiles(di.Parent.FullName, "*.dwg", SearchOption.TopDirectoryOnly);
            List<string> listFile = new List<string>();

            foreach (string file in filesNames)
            {
                listFile.Add(Path.GetFileNameWithoutExtension(file));
            }

            shortName = Rename(ref listFile, shortName);

            return shortName;
        }

        public static string Rename(ref List<string> names, string input)
        {
            string pattern = @"\(\d{1,3}\)$";

            if (names.Contains(input) == false)
            {
                names.Add(input);
                return input;
            }
            else
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(input, pattern))
                {
                    int num = Convert.ToInt16(System.Text.RegularExpressions.Regex.Match(input, pattern).Value.Replace("(", "").Replace(")", "")) + 1;
                    return Rename(ref names, input.Substring(0, input.Length - 3) + "(" + num + ")");
                }
                else
                {
                    return Rename(ref names, input + "(1)");
                }
            }
        }

    }
}
