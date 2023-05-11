using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

//creates a list of all CB specs from the PD folder
namespace Cardboard_Report
{
    public class Class1
    {
        //request if this is a master list or specific list
        //if this is a specific list
        //request a folder to search for CB specs under
        //use LINQ to find all dwgs under a cardboard folder in the requested folder
        //if this is a master list
        //use LINQ to find all dwgs under a cardboard folder in the PD folder
        //load the side DB of all dwgs and read through the blkrefs named cardboard spec
        //load the CB spec data to a csv, either to a temp folder if a specific file or to a particular spot
        // for a master list

        [CommandMethod("CBReport")]
        static public void ProcessCB()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;            

            //ask if a master file or a specific search
            PromptKeywordOptions pko = new PromptKeywordOptions("Master file or particular folder?");
            pko.Keywords.Add("Master");
            pko.Keywords.Add("Folder");
            pko.AllowArbitraryInput = false;

            bool master = false;

            ProgressMeter pm = new ProgressMeter();

            PromptResult pkeyRes = ed.GetKeywords(pko);
            if(pkeyRes.Status == PromptStatus.OK)
            {
                if(pkeyRes.StringResult.ToUpper() == "MASTER")
                { master = true; }
                else if (pkeyRes.StringResult.ToUpper() == "FOLDER")
                { master = false; }

                //get a folder to be searched
                string searchFolder = "";
                if(master == true)
                { searchFolder = @"Y:\Product Development\Style Specifications"; }
                else
                {
                    FolderBrowserDialog folder = new FolderBrowserDialog();
                    folder.SelectedPath = @"Y:\Product Development\Style Specifications";
                    folder.RootFolder = Environment.SpecialFolder.Desktop;

                   if(folder.ShowDialog() == DialogResult.OK)
                   { searchFolder = folder.SelectedPath; }
                }

                //look through the folder to find dwgs inside of Cardboard folders
                IEnumerable<FileInfo> CBfiles = LINQ(searchFolder);

                //choose a filelocation based on first choice
                string saveLoc = "";
                if(master == true)
                { saveLoc = @"Y:\Product Development\Standards\Cardboard List.csv"; }
                else
                { saveLoc = @"C:\temp\drawings\Cardboard list(" + DateTime.Now.ToShortDateString() + ").csv"; }

                //if data already exists then we need to clean it out
                //back up the last data and set up for new spot
                if(File.Exists(saveLoc))
                { File.Delete(saveLoc);}

                //loop through the files and act on them
                //set up a progress meter
                pm.Start("Writing data to file");
                pm.SetLimit(CBfiles.Count());
                foreach(FileInfo fi in CBfiles)
                {
                    if (hasCBfolder(fi.FullName))
                    { writeToCSV(fi, saveLoc, ed); }

                    //sleep to not waste CPU cycles
                    System.Threading.Thread.Sleep(5);
                    Autodesk.AutoCAD.ApplicationServices.Application.UpdateScreen();
                    pm.MeterProgress();
                }
                pm.Stop();
                ed.WriteMessage("\nFinished writing to:" + saveLoc);
            }
        }

        private static void writeToCSV(FileInfo fi, string saveLoc, Editor ed)
        {
            //if file is an autoCad file
            if (Path.GetExtension(fi.FullName) == ".dwg")
            {
                Database sideDb = new Database(false, true);
                using (sideDb)
                {
                    try { sideDb.ReadDwgFile(fi.FullName, FileShare.Read, false, ""); }
                    catch (System.Exception)
                    {
                        ed.WriteMessage("\nUnable to read drawing file.");
                        return;
                    }

                    using (Transaction tr = sideDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = sideDb.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                        //list all the block refs from model/paperspace
                        IEnumerable<BlockReference> blkRefList = getModalSpaceBlockRefs(sideDb, bt, btr);
                        IEnumerable<BlockReference> blkRefList2 = GetPaperSpaceBlockReferences(sideDb);

                        //check all blkrefs in database and write out all pertinant data
                        foreach (BlockReference blkRef in blkRefList.Concat(blkRefList2))
                        { writeToFile(blkRef, saveLoc, tr, fi.FullName); }
                    }
                }
            }
            else
            {
                if(Path.GetFileNameWithoutExtension(fi.Name).ToUpper() == "NO CARDBOARD" ||
                    Path.GetFileNameWithoutExtension(fi.Name).ToUpper() == "NOT APPLICABLE")
                {
                    //write to the file
                    StringBuilder attOut = new StringBuilder();
                    attOut.AppendFormat("{0},{1},{2}",
                        GetStyleID(fi.FullName),
                        "NA",
                        fi.LastWriteTime);
                    attOut.AppendLine();
                    File.AppendAllText(saveLoc, attOut.ToString());
                }
            }
        }

        //pass a blockref to method, search through for relevent data and write it to a file
        private static void writeToFile(BlockReference blkRef, string saveLoc, Transaction tr, string fileName)
        {
            //try to verify the block name but be sure to account for erased blockrefs
            //because this is a side base this should never be an issue...
            string blockName = "";
            try { blockName = blkRef.Name; }
            catch (System.Exception e)
            { return; }

            if(blockName == "CardboardForm")
            {
                AttributeCollection atts = blkRef.AttributeCollection;

                //try to build a string from the data
                //StyleID | partName | Date | W | L | thickness | grain |Quantity |description?
                string styleId = GetStyleID(fileName); //try to pull from the folder name
                string date = "No Date";
                string itemNum = "No Name";
                string length = "-";
                string width = "-";
                string thickness = "-";
                string grain = "-";
                string quantity = "-";
                string desc = "";

                //look for certain tags
                foreach(ObjectId id in atts)
                {
                    AttributeReference ar = tr.GetObject(id, OpenMode.ForRead) as AttributeReference;
                    if(ar != null && ar.TextString != "")
                    {
                        switch(ar.Tag.ToUpper())
                        {
                            case "CS_ITEMNUM":
                                itemNum = ar.TextString;
                                break;
                            case "CS_DATE":
                                date = ar.TextString;
                                break;
                            case "CS_LENGTH":
                                length = ar.TextString;
                                break;
                            case "CS_WIDTH":
                                width = ar.TextString;
                                break;
                            case "CS_THICKNESS":
                                thickness = ar.TextString;
                                break;
                            case "CS_DIR":
                                grain = ar.TextString;
                                break;
                            case "CS_QUANTITY":
                                quantity = ar.TextString;
                                break;
                            case "CS_ITEMDESC":
                                desc = ar.TextString;
                                break;
                        }
                    }                    
                }
                //arrange all data into a string and save it to the file
                StringBuilder attOut = new StringBuilder();
                attOut.AppendFormat("{0},{1},{2},{3},{4},{5},{6},{7}",
                    styleId, itemNum, date, width, length, thickness, grain, quantity);
                attOut.AppendLine();
                File.AppendAllText(saveLoc, attOut.ToString());
            }

        }

        private static IEnumerable<FileInfo> LINQ(string searchFolder)
        {
            DirectoryInfo dir = new DirectoryInfo(searchFolder);
            IEnumerable<FileInfo> fileList = dir.GetFiles("*.*", SearchOption.AllDirectories);

            //filter for dwgs
            fileList =
                from file in fileList
                where file.Extension == ".dwg" ||
                file.Extension == ".txt"
                orderby file.DirectoryName
                select file;
            
            return fileList;
        }

        //grab all blkrefs from model space
        private static IEnumerable<BlockReference> getModalSpaceBlockRefs(Database db, BlockTable bt, BlockTableRecord btr)
        {
            Transaction tr = db.TransactionManager.TopTransaction;
            if (tr == null)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NotTopTransaction);

            foreach (ObjectId msId in btr)
            {
                if (msId.ObjectClass.DxfName.ToUpper() == "INSERT")
                {
                    BlockReference blkRef = tr.GetObject(msId, OpenMode.ForRead) as BlockReference;
                    yield return blkRef;
                }
            }
        }

        //grab all blkrefs from paper spaces
        private static IEnumerable<BlockReference> GetPaperSpaceBlockReferences(Database db)
        {
            Transaction tr = db.TransactionManager.TopTransaction;
            if (tr == null)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NotTopTransaction);

            RXClass rxc = RXClass.GetClass(typeof(BlockReference));
            DBDictionary layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            foreach (var entry in layouts)
            {
                if (entry.Key != "Model")
                {
                    Layout lay = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        if (id.ObjectClass == rxc)
                        {
                            yield return (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        }
                    }
                }
            }
        }

        public static string GetStyleID(string pathName)
        {
            string styleID = "ID not Found";
            //string pathName = Path.GetDirectoryName(doc.Name);
            string[] styleIDparts = pathName.Split('\\');
            styleID = styleIDparts[styleIDparts.Length - 3];
            return styleID;
        }

        public static bool hasCBfolder(string pathName)
        {
            string styleID = "ID not Found";
            //string pathName = Path.GetDirectoryName(doc.Name);
            string[] styleIDparts = pathName.Split('\\');
            styleID = styleIDparts[styleIDparts.Length - 2];

            if (styleID.ToUpper() == "CARDBOARD" || styleID.ToUpper() == "CARD BOARD")
                return true;
            else
                return false;
        }
    }
}
