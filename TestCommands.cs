using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace LogiK3D.Piping
{
    public class TestCommands
    {
        [CommandMethod("TEST_BLOCK_GEN")]
        public void TestBlockGeneration()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptKeywordOptions pko = new PromptKeywordOptions("\nQuel bloc tester ? [Coude/Tee/Reduction/Bride/Vanne]: ");
            pko.Keywords.Add("Coude");
            pko.Keywords.Add("Tee");
            pko.Keywords.Add("Reduction");
            pko.Keywords.Add("Bride");
            pko.Keywords.Add("Vanne");
            pko.AllowNone = false;

            PromptResult pr = ed.GetKeywords(pko);
            if (pr.Status != PromptStatus.OK) return;

            PromptPointResult ppr = ed.GetPoint("\nPoint d'insertion : ");
            if (ppr.Status != PromptStatus.OK) return;

            PipingGenerator generator = new PipingGenerator();
            ObjectId blockId = ObjectId.Null;
            string blockName = "TEST_" + pr.StringResult + "_" + DateTime.Now.Ticks;

            try
            {
                switch (pr.StringResult)
                {
                    case "Coude":
                        blockId = generator.GetOrCreateElbow(114.3, 152.0, 90.0, blockName);
                        break;
                    case "Tee":
                        blockId = generator.GetOrCreateTee(114.3, 114.3, 200.0, 100.0, blockName);
                        break;
                    case "Reduction":
                        blockId = generator.GetOrCreateReducer(114.3, 88.9, 150.0, blockName);
                        break;
                    case "Bride":
                        blockId = generator.GetOrCreateFlange(114.3, 220.0, 16.0, 50.0, blockName);
                        break;
                    case "Vanne":
                        blockId = generator.GetOrCreateValve(250.0, 220.0, 150.0, false, blockName);
                        break;
                }

                if (blockId != ObjectId.Null)
                {
                    generator.InsertBlockReference(blockId, ppr.Value);
                    ed.WriteMessage($"\nBloc {pr.StringResult} généré et inséré avec succès !");
                }
                else
                {
                    ed.WriteMessage($"\nÉchec de la génération du bloc {pr.StringResult}.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nERREUR FATALE lors du test : {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}