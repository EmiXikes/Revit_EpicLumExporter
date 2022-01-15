using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using static DEL_acadltlib_EM.FileIO;
using System.Numerics;
using System.Drawing;
using Autodesk.AutoCAD.Geometry;
using System.Windows;
using EpicLumExporter.UI.ViewModel;
using EpicLumExporter.UI.View;

namespace EpicLumExporter
{
    public class ExportLums
    {
        static string pathFamilyNamesList = System.IO.Path.Combine("C:\\Epic\\Revit", "ELI_Fams.txt");
        static string pathLumDwgExportData = System.IO.Path.Combine("C:\\Epic\\Revit", "ELI_DwgLumData.txt");
        static string pathLumDwgSummaryData = System.IO.Path.Combine("C:\\Epic\\Revit", "ELI_DwgLumSummary.txt");
        static string pathLumInfoBlockData = System.IO.Path.Combine("C:\\Epic\\Revit", "ELI_DwgLumInfoBlcks.txt");
        static string pathLumOrigins = System.IO.Path.Combine("C:\\Epic\\Revit", "ELI_DwgLumOrigins.txt");

        [CommandMethod("LUMEXPORT")]
        public static void ExportL()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = ThisDrawing.Database;
            Editor ed = ThisDrawing.Editor;

            List<LumDataItem> LumData = new List<LumDataItem>();
            List<LumDataItem2> LumData2 = new List<LumDataItem2>();
            List<UniqueLumDataItem> LumUniqueData = new List<UniqueLumDataItem>();
            List<LumInfoBlock> LumInfoBlocks = new List<LumInfoBlock>();

            Vector2 LumInfoOrigins = new Vector2();

            using (Autodesk.AutoCAD.DatabaseServices.Transaction trans = db.TransactionManager.StartTransaction())
            {
                //db.ResolveXrefs(true, false);
                #region XREF STUFF

                //XrefGraph xg = db.GetHostDwgXrefGraph(false);
                //GraphNode xrefRoot = xg.RootNode;


                //List<string> xRefs = new List<string>();   
 
                //for (int o = 0; o < xrefRoot.NumOut; o++)
                //{
                //    XrefGraphNode child = xrefRoot.Out(o) as XrefGraphNode;
                //    if (child.XrefStatus == XrefStatus.Resolved)
                //    {
                //        xRefs.Add(child.Name);
                //    }
                //}

                //var vm = new XrefSelectorViewModel();
                //var xrefSelectorWin = new Window
                //{
                //    Content = new XrefSelector(),
                //    DataContext = vm,
                //    Width = 260,
                //    Height = 300,
                //    WindowStartupLocation = WindowStartupLocation.CenterScreen
                //};
                //vm.OnRequestClose += (s, e) => xrefSelectorWin.Close();
                //vm.xrefs = xRefs;
                //xrefSelectorWin.ShowDialog();

                //string selectedXrefName = xRefs[vm.selectedIndex];
                //Database dbXref = new Database();
                //for (int o = 0; o < xrefRoot.NumOut; o++)
                //{
                //    XrefGraphNode child = xrefRoot.Out(o) as XrefGraphNode;
                //    if (child.XrefStatus == XrefStatus.Resolved)
                //    {
                //        if (child.Name == selectedXrefName)
                //        {
                //            dbXref = child.Database;
                //        }
                //    }
                //}

                //var blockTableXref = (BlockTable)trans.GetObject(dbXref.BlockTableId, OpenMode.ForRead);
                //var modelSpaceXref = (BlockTableRecord)trans.GetObject(blockTableXref[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                //foreach(var eXref in modelSpaceXref)
                //{
                //    DBObject dbObj = trans.GetObject(eXref, OpenMode.ForRead);
                //    if (dbObj is BlockReference)
                //    {
                //        BlockReference blckRef = (BlockReference)dbObj;
                //        BlockTableRecord blckRecord = (BlockTableRecord)trans.GetObject(blckRef.DynamicBlockTableRecord, OpenMode.ForRead);
                //    }

                //}
                #endregion

                // open the block table which contains all the BlockTableRecords (block definitions and spaces)
                var blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (var E in modelSpace)
                {
                    DBObject dbObj = trans.GetObject(E, OpenMode.ForRead);

                    // Iterate through all block references in model space
                    if (dbObj is BlockReference)
                    {
                        BlockReference blckRef = (BlockReference)dbObj;
                        BlockTableRecord blckRecord = (BlockTableRecord)trans.GetObject(blckRef.DynamicBlockTableRecord, OpenMode.ForRead);
                        
                        // InfoBlocks
                        if (blckRecord.Name == "LumInfoBlock")
                        {
                            LumInfoBlock infoBlock = new LumInfoBlock();

                            foreach (ObjectId ID in blckRef.AttributeCollection)
                            {
                                var attRef = (AttributeReference)trans.GetObject(ID, OpenMode.ForRead);
                                if (attRef.Tag == "ELGRUPA")
                                {
                                    infoBlock.attr_ELGRUPA = attRef.TextString;
                                }
                                else if (attRef.Tag == "INFO1")
                                {
                                    infoBlock.attr_INFO1 = attRef.TextString;
                                }
                                else if(attRef.Tag == "INFO2")
                                {
                                    infoBlock.attr_INFO2 = attRef.TextString;
                                }
                                else if (attRef.Tag == "INFO3")
                                {
                                    infoBlock.attr_INFO3 = attRef.TextString;
                                }
                                else if (attRef.Tag == "COORDINATESTOPRIGHT")
                                {
                                    float X = (float)blckRef.Position.X + float.Parse(attRef.TextString.Split(';')[0], System.Globalization.CultureInfo.InvariantCulture);
                                    float Y = (float)blckRef.Position.Y + float.Parse(attRef.TextString.Split(';')[1], System.Globalization.CultureInfo.InvariantCulture); ;
                                    infoBlock.PointB = new Vector2(X, Y);
                                }
                            }
                            infoBlock.PointA = new Vector2((float)blckRef.Position.X, (float)blckRef.Position.Y);

                            LumInfoBlocks.Add(infoBlock);
                            continue;
                        }

                        if(blckRef.Name == "LumInfoOrigins" || blckRecord.Name == "LumInfoOrigins")
                        {
                            LumInfoOrigins = new Vector2((float)blckRef.Position.X, (float)blckRef.Position.Y);
                            continue;
                        }

                        // DLX luminaire table data block
                        if (blckRef.Layer.Contains("LUMKEY"))
                        {
                            bool dataBegins = false;
                            LumUniqueData = new List<UniqueLumDataItem>();
                            Table LumTable = (Table)blckRef;

                            for (int Row = 0; Row < LumTable.Rows.Count; Row++)
                            {
                                if (LumTable.Cells[Row, 0].Value.ToString() == "Index")
                                {
                                    dataBegins = true;
                                    continue;
                                }

                                if (dataBegins)
                                {
                                    LumUniqueData.Add(
                                        new UniqueLumDataItem
                                        {
                                            lumIndex = ReadFromCellExplodeText(LumTable, Row, "Index"),
                                            manufacturer = ReadFromCellExplodeText(LumTable, Row, "Manufacturer"),
                                            luminaireModelName = ReadFromCellExplodeText(LumTable, Row, "Article name"),
                                            quantity = ReadFromCellExplodeText(LumTable, Row, "Quantity")

                                            //manufacturer = LumTable.Cells[Row, HeaderColIndex(LumTable, "Manufacturer")].TextString.ToString(),
                                            //luminaireModelName = LumTable.Cells[Row, HeaderColIndex(LumTable, "Article name")].TextString.ToString(),
                                            //quantity = LumTable.Cells[Row, HeaderColIndex(LumTable, "Quantity")].TextString.ToString()
                                        });
                                }
                            }
                            continue;
                        }

                        // Skipping the loop for any other blocks that do not meet the requirements
                        if (!blckRef.Layer.StartsWith("DLX") || !blckRef.Layer.Contains("LUM")) { continue; }

                        // Luminaire blocks
                            // Getting index and model name
                        string lumindex = blckRef.Layer.PadRight(6).Split(' ')[1].Trim();
                        var lumTableDataItem = LumUniqueData.FirstOrDefault(L => L.lumIndex == lumindex);

                        string LumModelName;
                        string LumManufacturer;

                        if (lumTableDataItem == null)
                        {
                            System.Windows.Forms.MessageBox.Show(
                                "Kļūdains Dialux export fails. \n" +
                                "Nesakrīt gaismekļu tabula ar izvietotajiem plānā. \n" +
                                "Gaismekļu identifikācijai izmantot Index nr. \n"+
                                "Lūdzu pārbaudīt Dialux failu un veikt atkārtotu export.",
                                "ERROR",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            LumModelName = lumindex;
                            LumManufacturer = "-";
                            //return;
                        } else
                        {
                            LumModelName = LumUniqueData.FirstOrDefault(L => L.lumIndex == lumindex).luminaireModelName;
                            LumManufacturer = LumUniqueData.FirstOrDefault(L => L.lumIndex == lumindex).manufacturer;
                        }

                            // Adding Luminaire info to collection
                            // First entry
                        if (LumData.Count == 0)
                        {
                            LumData.Add(
                                new LumDataItem
                                {
                                    luminaireModelName = LumModelName,
                                    AcadBlckItem = blckRef
                                });
                            LumData2.Add(
                                new LumDataItem2
                                {
                                    LumModelName = LumModelName,
                                    LumManufacturer = LumManufacturer,
                                    Location = new Vector3()
                                    {
                                        X = (float)blckRef.Position.X,
                                        Y = (float)blckRef.Position.Y,
                                        Z = (float)blckRef.Position.Z
                                    },
                                    Rotation = blckRef.Rotation,
                                });
                        }
                        else
                        {
                            // Next entries
                            // Adding only one entry per given corrdinates (since DLX export creates multiple blocks in same references)
                            if (LumData.FirstOrDefault(L => L.AcadBlckItem.Position.ToString() == blckRef.Position.ToString()) == null)
                            {
                                LumData.Add(
                                    new LumDataItem
                                    {
                                        luminaireModelName = LumModelName,
                                        AcadBlckItem = blckRef
                                    });
                                LumData2.Add(
                                      new LumDataItem2
                                      {
                                          LumModelName = LumModelName,
                                          LumManufacturer = LumManufacturer,
                                          Location = new Vector3()
                                          {
                                              X = (float)blckRef.Position.X,
                                              Y = (float)blckRef.Position.Y,
                                              Z = (float)blckRef.Position.Z
                                          },
                                          Rotation = blckRef.Rotation,
                                      });
                                //Debug.Print(string.Format("Type: {0} [{1}]", dbObj.GetType().ToString(), dbObj.Handle));
                                Debug.Print(string.Format("Bref: {0} [{1}] <{2}>", blckRef.Name, blckRef.Layer, blckRef.Position.ToString()));
                            }
                        }

                    }
                }



                trans.Commit();
            }

            Debug.Print(string.Format("Total count: {0}", LumData.Count));

            bool isQuantityOK = true;
            string quantityErrorInfo = "";

            var gr = LumData.GroupBy(x => x.AcadBlckItem.Layer).ToList();
            foreach (var G in gr)
            {
                Debug.Print(string.Format("Type {0}:  {1}   [{2}]", G.Key, G.Count(), G.ToList()[0].luminaireModelName));
                ed.WriteMessage(string.Format("\nType {0}:  {1}   [{2}]", G.Key, G.Count(), G.ToList()[0].luminaireModelName));

                var tableItem = LumUniqueData.FirstOrDefault(L => L.luminaireModelName == G.ToList()[0].luminaireModelName);
                string gQuantity = G.Count().ToString().Trim();

                if (gQuantity != tableItem.quantity.Trim())
                {
                    isQuantityOK = false;
                    quantityErrorInfo += "[ " + tableItem.lumIndex + " ]  " + tableItem.luminaireModelName + "\n";
                }
            }
            if (!isQuantityOK)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Kļūdains Dialux export fails. \n" +
                    "Nesakrīt gaismekļu skaits tabulā ar izvietotajiem plānā. \n" +
                    "Lūdzu pārbaudīt.\n\n" +
                    "Kļūdainās pozīcijas:\n" + quantityErrorInfo,
                    "NESAKRITĪBA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            List<LumInfoRect> rectList = new List<LumInfoRect>();
            DrawRectangle(
                            new Point2d(0, 0),
                            new Point2d(1250,1250)
                            );

            foreach (var item in LumInfoBlocks)
            {
                LumInfoRect LumInfoRect = new LumInfoRect()
                {
                    Rect = new Rectangle(
                    (int)item.PointA.X,
                    (int)item.PointA.Y,
                    Math.Abs((int)item.PointA.X - (int)item.PointB.X),
                    Math.Abs((int)item.PointA.Y - (int)item.PointB.Y)
                    ),
                    attr_ELGRUPA = item.attr_ELGRUPA,
                    attr_INFO1 = item.attr_INFO1,
                    attr_INFO2 = item.attr_INFO2,
                    attr_INFO3 = item.attr_INFO3,
                };

                rectList.Add(LumInfoRect);

                DrawRectangle(
                new Point2d(item.PointA.X, item.PointA.Y),
                new Point2d(item.PointA.X+20, item.PointA.Y+20)
                );

                DrawRectangle(
                new Point2d(item.PointB.X, item.PointB.Y),
                new Point2d(item.PointB.X + 20, item.PointB.Y + 20)
                );

                DrawRectangle(
                    new Point2d(LumInfoRect.Rect.X, LumInfoRect.Rect.Y),
                    new Point2d(LumInfoRect.Rect.X + LumInfoRect.Rect.Width, LumInfoRect.Rect.Y + LumInfoRect.Rect.Height)
                    );

            }

            rectList.OrderBy(x=>x.Rect.Width * x.Rect.Height).Reverse();

            foreach(var rect in rectList)
            {
                foreach (var Lum in LumData2)
                {
                    System.Drawing.Point blckCoords = new System.Drawing.Point((int)Lum.Location.X, (int)Lum.Location.Y);
                    if (rect.Rect.Contains(blckCoords))
                    {
                        Lum.attr_ELGRUPA = rect.attr_ELGRUPA;
                    }
                }

            }
            


            SaveObjToFile(pathLumOrigins, LumInfoOrigins);
            SaveObjToFile(pathLumDwgExportData, LumData2);
            SaveObjToFile(pathLumDwgSummaryData, LumUniqueData);
            SaveObjToFile(pathLumInfoBlockData, LumInfoBlocks);

        }

        private static string ReadFromCellExplodeText(Table LumTable, int Row, string columnName)
        {

            var C = LumTable.Cells[Row, HeaderColIndex(LumTable, columnName)];
            if (C.Value != null)
            {
                MText mText = new MText();
                mText.Contents = C.TextString;
                DBObjectCollection dbs = new DBObjectCollection();
                mText.Explode(dbs);
                string s = "";
                foreach (Entity en in dbs)
                {
                    DBText txt = en as DBText;
                    if (txt != null)
                    {
                        s += txt.TextString;
                    }
                }
                mText.Dispose();
                dbs.Dispose();
                return s;
            }

            return "";
            //return LumTable.Cells[Row, HeaderColIndex(LumTable, columnName)].TextString.ToString();
        }
        public static int HeaderColIndex(Table refTable, string headerName)
        {
            for (int Row = 0; Row < refTable.Rows.Count; Row++)
            {
                for (int Col = 0; Col < refTable.Columns.Count; Col++)
                {
                    var C = refTable.Cells[Row, Col];
                    if (C.Value != null)
                    {
                        MText mText = new MText();
                        mText.Contents = C.TextString;
                        DBObjectCollection dbs = new DBObjectCollection();
                        mText.Explode(dbs);
                        string s = "";
                        foreach (Entity en in dbs)
                        {
                            DBText txt = en as DBText;
                            if (txt != null)
                            {
                                s += txt.TextString;
                            }
                        }
                        mText.Dispose();
                        dbs.Dispose();

                        if (headerName == s)
                        {
                            return Col;
                        }
                    }

                }
            }
            return -1;
        }

        public static void DrawRectangle(Point2d pt1, Point2d pt2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = ThisDrawing.Database;
            Editor ed = ThisDrawing.Editor;



            //var ppr = ed.GetPoint("\nFirst corner: ");
            //if (ppr.Status != PromptStatus.OK)
            //    return;
            //var pt1 = ppr.Value;

            //ppr = ed.GetCorner("\nOpposite corner: ", pt1);
            //if (ppr.Status != PromptStatus.OK)
            //    return;
            //var pt2 = ppr.Value;

            double x1 = pt1.X;
            double y1 = pt1.Y;
            double x2 = pt2.X;
            double y2 = pt2.Y;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pline = new Polyline(4);
                pline.AddVertexAt(0, new Point2d(x1, y1), 0.0, 0.0, 0.0);
                pline.AddVertexAt(1, new Point2d(x2, y1), 0.0, 0.0, 0.0);
                pline.AddVertexAt(2, new Point2d(x2, y2), 0.0, 0.0, 0.0);
                pline.AddVertexAt(3, new Point2d(x1, y2), 0.0, 0.0, 0.0);
                pline.Closed = true;
                pline.TransformBy(ed.CurrentUserCoordinateSystem);
                var curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                curSpace.AppendEntity(pline);
                tr.AddNewlyCreatedDBObject(pline, true);
                tr.Commit();
            }
        }

    }

    public class LumInfoBlock
    {
        public Vector2 PointA;
        public Vector2 PointB;
        public string attr_ELGRUPA;
        public string attr_INFO1;
        public string attr_INFO2;
        public string attr_INFO3;

    }
    public class LumInfoRect
    {
        public Rectangle Rect;
        public string attr_ELGRUPA;
        public string attr_INFO1;
        public string attr_INFO2;
        public string attr_INFO3;

    }

    public class LumDataItem2
    {
        public string LumModelName;
        public string LumManufacturer;
        public Vector3 Location;
        public double Rotation;
        public string attr_ELGRUPA;
        //public BlockReference AcadBlckItem;
    }
    public class LumDataItem
    {
        public string luminaireModelName;
        public BlockReference AcadBlckItem;
    }

    public class UniqueLumDataItem
    {
        public string lumIndex;
        public string manufacturer;
        public string luminaireModelName;
        public string itemNumber;
        public string quantity;
    };



}
