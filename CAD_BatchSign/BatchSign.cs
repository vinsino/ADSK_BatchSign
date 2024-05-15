using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using DocumentProcessing;
using System.Collections.Generic;
using System.IO;

[assembly: CommandClass(typeof(CAD_BatchSign.BatchSign))]

namespace CAD_BatchSign
{
    public class BatchSign
    {
        [CommandMethod("SIGNMYNAME")]
        public void SignMyName()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptKeywordOptions pko = new PromptKeywordOptions("\nChose one option [Paper/Model]: ", "Paper Model");
            pko.Keywords.Default = "Paper";
            PromptResult pres = ed.GetKeywords(pko);
            if (pres.Status != PromptStatus.OK)
                return;

            ExcelReader excelReader = new ExcelReader();
            PdfReader pdfReader = new PdfReader();
            var data = excelReader.GetDrawingSigns();


            List<Point3d> name_alignment = new List<Point3d> {
                new Point3d(75.3399, 12.15, 0),
                new Point3d(75.3399, 11.15, 0),
                new Point3d(80.779186, 13.1, 0),
                new Point3d(80.779186, 12.15, 0),
                new Point3d(80.779186, 11.15, 0),
            };

            List<Point3d> date_alignment = new List<Point3d> {
                new Point3d(77.0001, 12.15, 0),
                new Point3d(77.0001, 11.15, 0),
                new Point3d(82, 12.95, 0),
                new Point3d(82, 12.15, 0),
                new Point3d(82, 11.15, 0),
            };

            List<Point3d> bureau_alignment = new List<Point3d> {
                new Point3d(76.201, 19.125, 0),
                new Point3d(77.392, 19.125, 0),
                new Point3d(78.448, 19.125, 0),
            };

            foreach (var item in data)
            {
                string dwg_path = Path.Combine(Path.GetDirectoryName(excelReader.FilePath), $@"dwg\{item.FileName}.dwg");

                if (!File.Exists(dwg_path)) continue;

                string pdf_path = Path.Combine(Path.GetDirectoryName(excelReader.FilePath), $@"pdf\{item.FileName}.pdf");
                item.SignDates = pdfReader.ExtractDateText(pdf_path);

                using (Database db = new Database(false, true))
                {
                    db.ReadDwgFile(dwg_path, FileOpenMode.OpenForReadAndWriteNoShare, false, "");
                    db.CloseInput(true);

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {

                        BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;

                        BlockTableRecord acBlkTblRec;

                        if (pres.StringResult == "Paper")
                            acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                        else
                            acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        TextStyleTable textStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                        for (int i = 0; i != item.SignNames.Count; i++)
                        {

                            // name
                            using (DBText acDBText = new DBText())
                            {
                                acDBText.SetDatabaseDefaults(db);

                                acDBText.TextString = item.SignNames[i];
                                acDBText.Justify = AttachmentPoint.MiddleCenter;
                                acDBText.AlignmentPoint = name_alignment[i];
                                acDBText.Height = 0.35;
                                acDBText.WidthFactor = 0.85;
                                acDBText.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);

                                if (layerTable.Has("變更欄-簽名") == false)
                                {
                                    using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                                    {
                                        // Assign the layer the ACI color 3 and a name
                                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                                        acLyrTblRec.Name = "變更欄-簽名";

                                        // Upgrade the Layer table for write
                                        tr.GetObject(db.LayerTableId, OpenMode.ForWrite);

                                        // Append the new layer to the Layer table and the transaction
                                        layerTable.Add(acLyrTblRec);
                                        tr.AddNewlyCreatedDBObject(acLyrTblRec, true);
                                    }
                                }

                                acDBText.LayerId = layerTable["變更欄-簽名"];

                                acDBText.TextStyleId = textStyleTable["Sino241"];

                                Database prev = HostApplicationServices.WorkingDatabase;
                                HostApplicationServices.WorkingDatabase = db;
                                acDBText.AdjustAlignment(db);
                                HostApplicationServices.WorkingDatabase = prev;

                                acBlkTblRec.AppendEntity(acDBText);
                                tr.AddNewlyCreatedDBObject(acDBText, true);
                            }

                            // date
                            using (DBText acDBText = new DBText())
                            {
                                acDBText.SetDatabaseDefaults(db);

                                acDBText.TextString = item.SignDates[i];
                                acDBText.Justify = AttachmentPoint.MiddleCenter;
                                acDBText.AlignmentPoint = date_alignment[i];
                                acDBText.Height = 0.18;
                                acDBText.WidthFactor = 1;
                                acDBText.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                                acDBText.LayerId = layerTable["變更欄-簽名"];

                                if (textStyleTable.Has("Sino2") == false)
                                {
                                    using (TextStyleTableRecord actsTblRec = new TextStyleTableRecord())
                                    {
                                        actsTblRec.Name = "Sino2";

                                        textStyleTable.UpgradeOpen();

                                        textStyleTable.Add(actsTblRec);

                                        actsTblRec.PriorSize = 0.4;
                                        actsTblRec.XScale = 1;
                                        actsTblRec.Font = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor("微軟正黑體", false, false, 136, 39);

                                        tr.AddNewlyCreatedDBObject(actsTblRec, true);

                                    }
                                }

                                acDBText.TextStyleId = textStyleTable["Sino2"];


                                acDBText.LineWeight = LineWeight.ByLayer;

                                Database prev = HostApplicationServices.WorkingDatabase;
                                HostApplicationServices.WorkingDatabase = db;
                                acDBText.AdjustAlignment(db);
                                HostApplicationServices.WorkingDatabase = prev;

                                acBlkTblRec.AppendEntity(acDBText);
                                tr.AddNewlyCreatedDBObject(acDBText, true);
                            }

                            // Bureau date
                            if (i > 2 || string.IsNullOrEmpty(item.SignDates[i + 5])) continue;

                            using (DBText acDBText = new DBText())
                            {
                                acDBText.SetDatabaseDefaults(db);

                                acDBText.TextString = item.SignDates[i + 5];
                                acDBText.Justify = AttachmentPoint.MiddleCenter;
                                acDBText.AlignmentPoint = bureau_alignment[i];
                                acDBText.Height = 0.31;
                                acDBText.WidthFactor = 1;
                                acDBText.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                                acDBText.LayerId = layerTable["變更欄-簽名"];
                                acDBText.TextStyleId = textStyleTable["Sino2"];


                                acDBText.LineWeight = LineWeight.ByLayer;

                                Database prev = HostApplicationServices.WorkingDatabase;
                                HostApplicationServices.WorkingDatabase = db;
                                acDBText.AdjustAlignment(db);
                                HostApplicationServices.WorkingDatabase = prev;

                                acBlkTblRec.AppendEntity(acDBText);
                                tr.AddNewlyCreatedDBObject(acDBText, true);
                            }

                        }

                        tr.Commit();
                    }

                    db.SaveAs(dwg_path, DwgVersion.AC1027);

                    ed.WriteMessage($"\n{item.FileName} completed...");
                }

            }

        }
    }
}
