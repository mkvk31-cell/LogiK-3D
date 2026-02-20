using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace LogiK3D.Piping
{
    public class PipingGenerator
    {
        private void AddInvisibleAttributes(BlockTableRecord btr, Transaction tr, params string[] extraTags)
        {
            System.Collections.Generic.List<string> tags = new System.Collections.Generic.List<string> { "POSITION", "DESIGNATION", "DN", "PN", "ID", "NO LIGNE" };
            if (extraTags != null)
            {
                tags.AddRange(extraTags);
            }

            foreach (string tag in tags)
            {
                AttributeDefinition attrDef = new AttributeDefinition();
                attrDef.Position = Point3d.Origin;
                attrDef.Tag = tag;
                attrDef.Prompt = tag;
                attrDef.Invisible = true;
                btr.AppendEntity(attrDef);
                tr.AddNewlyCreatedDBObject(attrDef, true);
            }
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de coude (Elbow) dans la base de données AutoCAD.
        /// </summary>
        /// <param name="diameter">Diamètre extérieur du coude.</param>
        /// <param name="radius">Rayon de courbure du coude (à l'axe).</param>
        /// <param name="angle">Angle du coude en degrés (ex: 90).</param>
        /// <param name="blockName">Nom unique du bloc (ex: ELBOW_90_DN100_R152).</param>
        /// <returns>L'ObjectId du BlockTableRecord créé ou existant.</returns>
        public ObjectId GetOrCreateElbow(double diameter, double radius, double angle, string blockName)
        {
            // Récupération du document actif et de sa base de données
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            // Utilisation d'une transaction pour garantir l'intégrité de la base de données
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. Check de la BlockTable
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Si le bloc existe déjà, on retourne son ObjectId
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                // 2. Création du "Jig" ou Géométrie Temporaire
                // Le bloc n'existe pas, on doit le créer. On passe la BlockTable en mode écriture.
                bt.UpgradeOpen();

                // Création du nouveau BlockTableRecord
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                // Le bloc est créé à l'origine (0,0,0) par défaut

                // Conversion de l'angle en radians pour les calculs géométriques
                double angleRad = angle * (Math.PI / 180.0);

                // Création des objets temporaires pour le Sweep (Balayage)
                using (Circle profile = new Circle())
                using (Arc path = new Arc())
                {
                    // Définition du chemin (Arc)
                    path.Center = Point3d.Origin;
                    path.Normal = Vector3d.ZAxis; // L'arc est dessiné dans le plan XY
                    path.Radius = radius;
                    path.StartAngle = 0.0;
                    path.EndAngle = angleRad;

                    // Définition du profil (Cercle)
                    // Créé à l'origine avec normale Z, puis déplacé et orienté
                    profile.Center = Point3d.Origin;
                    profile.Normal = Vector3d.ZAxis;
                    profile.Radius = diameter / 2.0;
                    
                    // Déplacement au point de départ de l'arc et orientation (normale Y)
                    profile.TransformBy(Matrix3d.Displacement(new Vector3d(radius, 0, 0)));
                    profile.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, new Point3d(radius, 0, 0)));

                    // Création du solide 3D
                    Solid3d solidElbow = new Solid3d();
                    try
                    {
                        // Options de balayage par défaut
                        SweepOptionsBuilder sweepOpts = new SweepOptionsBuilder();
                        
                        // Génération du solide par balayage du profil le long du chemin
                        solidElbow.CreateSweptSolid(profile, path, sweepOpts.ToSweepOptions());

                        // Déplacer le solide pour que le point de départ (port 1) soit à l'origine (0,0,0)
                        solidElbow.TransformBy(Matrix3d.Displacement(Point3d.Origin - path.StartPoint));
                        
                        // Faire pivoter pour que la direction de départ soit l'axe X (au lieu de Y)
                        solidElbow.TransformBy(Matrix3d.Rotation(-Math.PI / 2.0, Vector3d.ZAxis, Point3d.Origin));

                        // Ajout du solide au BlockTableRecord
                        btr.AppendEntity(solidElbow);
                        tr.AddNewlyCreatedDBObject(solidElbow, true);
                    }
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage($"\nErreur lors de la création du solide balayé (Coude) : {ex.Message}");
                        solidElbow.Dispose();
                    }

                    // 3. Points d'accroche (Ports)
                    // Création des DBPoints aux extrémités pour faciliter l'accrochage (Osnap)
                    DBPoint startPort = new DBPoint(Point3d.Origin);
                    
                    Point3d endPt = (path.EndPoint - (path.StartPoint - Point3d.Origin)).TransformBy(Matrix3d.Rotation(-Math.PI / 2.0, Vector3d.ZAxis, Point3d.Origin));
                    DBPoint endPort = new DBPoint(endPt);
                    
                    double cutback = radius * Math.Tan(angleRad / 2.0);
                    DBPoint centerPort = new DBPoint(new Point3d(cutback, 0, 0));
                    
                    // Assignation au calque DEFPOINTS (calque non imprimable standard d'AutoCAD)
                    try { startPort.Layer = "DEFPOINTS"; } catch { }
                    try { endPort.Layer = "DEFPOINTS"; } catch { }
                    try { centerPort.Layer = "DEFPOINTS"; } catch { }

                    // Ajout des points au BlockTableRecord
                    btr.AppendEntity(startPort);
                    tr.AddNewlyCreatedDBObject(startPort, true);
                    btr.AppendEntity(endPort);
                    tr.AddNewlyCreatedDBObject(endPort, true);
                    btr.AppendEntity(centerPort);
                    tr.AddNewlyCreatedDBObject(centerPort, true);
                } // Les objets temporaires (profile, path) sont automatiquement Dispose() ici

                // Ajout des attributs invisibles
                AddInvisibleAttributes(btr, tr);

                // Validation de la transaction
                tr.Commit();
            }

            return blockId;
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de Té (Tee) dans la base de données AutoCAD.
        /// </summary>
        /// <param name="mainDiameter">Diamètre extérieur du tube principal.</param>
        /// <param name="branchDiameter">Diamètre extérieur de la dérivation.</param>
        /// <param name="length">Longueur totale du corps principal.</param>
        /// <param name="branchHeight">Hauteur de la dérivation depuis l'axe central.</param>
        /// <param name="blockName">Nom unique du bloc (ex: TEE_EGAL_DN100).</param>
        /// <returns>L'ObjectId du BlockTableRecord créé ou existant.</returns>
        public ObjectId GetOrCreateTee(double mainDiameter, double branchDiameter, double length, double branchHeight, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // Création du corps principal (Cylindre aligné sur l'axe X)
                Solid3d mainSolid = new Solid3d();
                try
                {
                    // CreateFrustum crée un cylindre le long de l'axe Z centré en 0,0,0
                    // Paramètres: height, xRadius, yRadius, topXRadius
                    mainSolid.CreateFrustum(length, mainDiameter / 2.0, mainDiameter / 2.0, mainDiameter / 2.0);
                    
                    // Rotation pour l'aligner sur l'axe X
                    mainSolid.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    // Création de la branche (Cylindre le long de l'axe Y)
                    Solid3d branchSolid = new Solid3d();
                    branchSolid.CreateFrustum(branchHeight, branchDiameter / 2.0, branchDiameter / 2.0, branchDiameter / 2.0);
                    
                    // Rotation pour l'aligner sur l'axe Y
                    branchSolid.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, Point3d.Origin)); 
                    
                    // Décalage pour que la base commence à l'axe (0,0,0) et aille jusqu'à branchHeight
                    branchSolid.TransformBy(Matrix3d.Displacement(new Vector3d(0, branchHeight / 2.0, 0)));

                    // Union booléenne : on ajoute la branche au corps principal
                    // ATTENTION: BooleanOperation détruit branchSolid. Ne pas utiliser de 'using' !
                    mainSolid.BooleanOperation(BooleanOperationType.BoolUnite, branchSolid);

                    // Déplacer le solide pour que le port 1 soit à l'origine (0,0,0)
                    mainSolid.TransformBy(Matrix3d.Displacement(new Vector3d(length / 2.0, 0, 0)));

                    btr.AppendEntity(mainSolid);
                    tr.AddNewlyCreatedDBObject(mainSolid, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du solide (Té) : {ex.Message}");
                    mainSolid.Dispose();
                }

                // Points d'accroche (Ports)
                DBPoint port1 = new DBPoint(Point3d.Origin);
                DBPoint port2 = new DBPoint(new Point3d(length, 0, 0));
                DBPoint port3 = new DBPoint(new Point3d(length / 2.0, branchHeight, 0));
                
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                try { port3.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);
                btr.AppendEntity(port3);
                tr.AddNewlyCreatedDBObject(port3, true);

                // Ajout des attributs invisibles
                AddInvisibleAttributes(btr, tr);

                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de Réduction (Reducer) concentrique.
        /// </summary>
        public ObjectId GetOrCreateReducer(double largeDiameter, double smallDiameter, double length, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                Solid3d reducerSolid = new Solid3d();
                try
                {
                    // Tronc de cône le long de l'axe Z
                    reducerSolid.CreateFrustum(length, smallDiameter / 2.0, smallDiameter / 2.0, largeDiameter / 2.0);
                    
                    // Déplacer pour que la base (grand diamètre) soit à Z=0 et le sommet à Z=length
                    reducerSolid.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, length / 2.0)));
                    
                    // Rotation pour l'aligner sur l'axe X (Z devient X)
                    reducerSolid.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    btr.AppendEntity(reducerSolid);
                    tr.AddNewlyCreatedDBObject(reducerSolid, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du solide (Réduction) : {ex.Message}");
                    reducerSolid.Dispose();
                }

                DBPoint port1 = new DBPoint(Point3d.Origin); // Côté large
                DBPoint port2 = new DBPoint(new Point3d(length, 0, 0));  // Côté réduit
                
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);

                AddInvisibleAttributes(btr, tr, "GRAND DN", "PETIT DN");

                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de Bride (Flange).
        /// </summary>
        public ObjectId GetOrCreateFlange(double pipeDiameter, double flangeDiameter, double thickness, double length, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                Solid3d flange = new Solid3d();
                try
                {
                    double neckLength = length - thickness;
                    if (neckLength > 0)
                    {
                        // Créer le collet d'abord (de Z=0 à Z=neckLength)
                        flange.CreateFrustum(neckLength, pipeDiameter / 2.0, pipeDiameter / 2.0, pipeDiameter / 2.0);
                        flange.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, neckLength / 2.0)));
                        
                        // Créer la face (de Z=neckLength à Z=length)
                        Solid3d face = new Solid3d();
                        face.CreateFrustum(thickness, flangeDiameter / 2.0, flangeDiameter / 2.0, flangeDiameter / 2.0);
                        // Léger overlap pour éviter les erreurs booléennes
                        face.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, neckLength + (thickness / 2.0) - 0.1)));
                        
                        flange.BooleanOperation(BooleanOperationType.BoolUnite, face);
                    }
                    else
                    {
                        flange.CreateFrustum(thickness, flangeDiameter / 2.0, flangeDiameter / 2.0, flangeDiameter / 2.0);
                        flange.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, thickness / 2.0)));
                    }

                    // Rotation pour aligner sur l'axe X
                    flange.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    btr.AppendEntity(flange);
                    tr.AddNewlyCreatedDBObject(flange, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du solide (Bride) : {ex.Message}");
                    flange.Dispose();
                }

                DBPoint port1 = new DBPoint(Point3d.Origin); // Côté soudure
                DBPoint port2 = new DBPoint(new Point3d(length, 0, 0)); // Côté face
                
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);

                AddInvisibleAttributes(btr, tr);

                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Importe un fichier DWG externe (ex: téléchargé depuis TraceParts) comme un bloc LogiK 3D.
        /// </summary>
        /// <param name="dwgFilePath">Chemin complet vers le fichier DWG téléchargé.</param>
        /// <param name="blockName">Nom à donner au bloc dans AutoCAD.</param>
        /// <returns>L'ObjectId du BlockTableRecord importé.</returns>
        public ObjectId ImportExternalDwgAsBlock(string dwgFilePath, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database destDb = doc.Database;
            ObjectId blockId = ObjectId.Null;

            if (!System.IO.File.Exists(dwgFilePath))
            {
                doc.Editor.WriteMessage($"\nErreur : Le fichier {dwgFilePath} est introuvable.");
                return blockId;
            }

            using (Transaction tr = destDb.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(destDb.BlockTableId, OpenMode.ForRead);
                
                // Si le bloc existe déjà, on le retourne
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                // Importer le DWG externe dans la base de données courante
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    destDb.Insert(blockName, sourceDb, true);
                }

                // Maintenant que le bloc est importé, on lui ajoute les attributs LogiK 3D
                bt.UpgradeOpen();
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForWrite);
                
                // Ajout des attributs invisibles (POSITION, DESIGNATION, DN, etc.)
                // On vérifie d'abord si les attributs n'existent pas déjà
                bool hasAttributes = false;
                foreach (ObjectId entId in btr)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent is AttributeDefinition)
                    {
                        hasAttributes = true;
                        break;
                    }
                }

                if (!hasAttributes)
                {
                    AddInvisibleAttributes(btr, tr);
                }

                blockId = bt[blockName];
                tr.Commit();
            }

            return blockId;
        }

        /// <summary>
        /// Insère une référence de bloc dans le ModelSpace.
        /// </summary>
        /// <param name="blockId">L'ObjectId du bloc à insérer.</param>
        /// <param name="position">La position d'insertion.</param>
        /// <param name="attributes">Dictionnaire des attributs à renseigner.</param>
        /// <returns>L'ObjectId de la BlockReference insérée.</returns>
        public ObjectId InsertBlockReference(ObjectId blockId, Point3d position, System.Collections.Generic.Dictionary<string, string> attributes = null)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId brefId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                BlockReference bref = new BlockReference(position, blockId);
                brefId = modelSpace.AppendEntity(bref);
                tr.AddNewlyCreatedDBObject(bref, true);

                // Ajout des attributs
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                foreach (ObjectId entId in btr)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent is AttributeDefinition attrDef)
                    {
                        AttributeReference attrRef = new AttributeReference();
                        attrRef.SetAttributeFromBlock(attrDef, bref.BlockTransform);
                        
                        if (attributes != null && attributes.ContainsKey(attrDef.Tag.ToUpper()))
                        {
                            attrRef.TextString = attributes[attrDef.Tag.ToUpper()];
                        }
                        
                        bref.AttributeCollection.AppendAttribute(attrRef);
                        tr.AddNewlyCreatedDBObject(attrRef, true);
                    }
                }

                tr.Commit();
            }

            return brefId;
        }

        /// <summary>
        /// Génère un bloc 3D de Vanne (Manuelle ou Pneumatique).
        /// </summary>
        public ObjectId GetOrCreateValve(double length, double flangeDiameter, double actuatorHeight, bool isPneumatic, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                Solid3d body = new Solid3d();
                try
                {
                    // Corps de la vanne (Cylindre)
                    body.CreateFrustum(length, flangeDiameter / 2.0, flangeDiameter / 2.0, flangeDiameter / 2.0);
                    body.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, length / 2.0)));
                    body.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    // Tige (Stem) - On la fait descendre un peu dans le corps pour éviter les erreurs booléennes
                    Solid3d stem = new Solid3d();
                    double stemHeight = actuatorHeight * 0.65;
                    stem.CreateFrustum(stemHeight, flangeDiameter * 0.15, flangeDiameter * 0.15, flangeDiameter * 0.15);
                    stem.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, stemHeight / 2.0)));
                    stem.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, Point3d.Origin)); // Aligné sur Y
                    stem.TransformBy(Matrix3d.Displacement(new Vector3d(length / 2.0, 0, 0))); // Centré sur le corps
                    body.BooleanOperation(BooleanOperationType.BoolUnite, stem);

                    // Actionneur (Manuel = levier, Pneumatique = gros cylindre)
                    Solid3d actuator = new Solid3d();
                    if (isPneumatic)
                    {
                        double actHeight = actuatorHeight * 0.4;
                        actuator.CreateFrustum(actHeight, flangeDiameter * 0.4, flangeDiameter * 0.4, flangeDiameter * 0.4);
                        actuator.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, actHeight / 2.0)));
                        actuator.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, Point3d.Origin));
                        // On le descend un peu pour qu'il fusionne bien avec la tige
                        actuator.TransformBy(Matrix3d.Displacement(new Vector3d(length / 2.0, actuatorHeight * 0.6, 0)));
                    }
                    else
                    {
                        // Levier manuel (Boîte)
                        actuator.CreateBox(length * 0.8, flangeDiameter * 0.1, flangeDiameter * 0.1);
                        actuator.TransformBy(Matrix3d.Displacement(new Vector3d(length / 2.0, actuatorHeight * 0.6, 0)));
                    }
                    body.BooleanOperation(BooleanOperationType.BoolUnite, actuator);

                    btr.AppendEntity(body);
                    tr.AddNewlyCreatedDBObject(body, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création de la vanne : {ex.Message}");
                    body.Dispose();
                }

                DBPoint port1 = new DBPoint(Point3d.Origin);
                DBPoint port2 = new DBPoint(new Point3d(length, 0, 0));
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);

                AddInvisibleAttributes(btr, tr, "TYPE VANNE", "ACTIONNEUR");

                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère un bloc 3D de Filtre Y.
        /// </summary>
        public ObjectId GetOrCreateYFilter(double length, double flangeDiameter, double filterHeight, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                Solid3d body = new Solid3d();
                try
                {
                    body.CreateFrustum(length, flangeDiameter / 2.0, flangeDiameter / 2.0, flangeDiameter / 2.0);
                    body.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, length / 2.0)));
                    body.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    // Branche du filtre Y (inclinée à 45 degrés vers le bas)
                    Solid3d branch = new Solid3d();
                    branch.CreateFrustum(filterHeight, flangeDiameter * 0.3, flangeDiameter * 0.3, flangeDiameter * 0.3);
                    branch.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, filterHeight / 2.0)));
                    branch.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, Point3d.Origin)); // Aligné sur Y
                    branch.TransformBy(Matrix3d.Rotation(-Math.PI / 4.0, Vector3d.ZAxis, Point3d.Origin)); // Rotation 45°
                    branch.TransformBy(Matrix3d.Displacement(new Vector3d(length / 2.0, 0, 0))); // Centrer sur le corps

                    body.BooleanOperation(BooleanOperationType.BoolUnite, branch);
                    btr.AppendEntity(body);
                    tr.AddNewlyCreatedDBObject(body, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du filtre Y : {ex.Message}");
                    body.Dispose();
                }

                DBPoint port1 = new DBPoint(Point3d.Origin);
                DBPoint port2 = new DBPoint(new Point3d(length, 0, 0));
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);

                AddInvisibleAttributes(btr, tr, "MAILLE");

                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère un bloc 3D de Débitmètre.
        /// </summary>
        public ObjectId GetOrCreateFlowmeter(double length, double flangeDiameter, double transmitterHeight, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                Solid3d body = new Solid3d();
                try
                {
                    body.CreateFrustum(length, flangeDiameter / 2.0, flangeDiameter / 2.0, flangeDiameter / 2.0);
                    body.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, length / 2.0)));
                    body.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    // Transmetteur (Boîte sur le dessus)
                    Solid3d transmitter = new Solid3d();
                    transmitter.CreateBox(flangeDiameter * 0.6, flangeDiameter * 0.6, transmitterHeight * 0.5);
                    // On le descend un peu pour qu'il fusionne bien avec le corps
                    transmitter.TransformBy(Matrix3d.Displacement(new Vector3d(length / 2.0, flangeDiameter / 2.0 + transmitterHeight * 0.2, 0)));
                    
                    body.BooleanOperation(BooleanOperationType.BoolUnite, transmitter);
                    btr.AppendEntity(body);
                    tr.AddNewlyCreatedDBObject(body, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du débitmètre : {ex.Message}");
                    body.Dispose();
                }

                DBPoint port1 = new DBPoint(Point3d.Origin);
                DBPoint port2 = new DBPoint(new Point3d(length, 0, 0));
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);

                AddInvisibleAttributes(btr, tr, "SIGNAL");

                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère un bloc 3D de Diaphragme (Plaque à orifice).
        /// </summary>
        public ObjectId GetOrCreateDiaphragm(double thickness, double outerDiameter, double tabHeight, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                Solid3d body = new Solid3d();
                try
                {
                    // Plaque principale
                    body.CreateFrustum(thickness, outerDiameter / 2.0, outerDiameter / 2.0, outerDiameter / 2.0);
                    body.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, thickness / 2.0)));
                    body.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    // Languette de lecture (Tab)
                    Solid3d tab = new Solid3d();
                    tab.CreateBox(thickness, outerDiameter * 0.2, tabHeight);
                    // On la descend un peu pour qu'elle fusionne bien avec le corps
                    tab.TransformBy(Matrix3d.Displacement(new Vector3d(thickness / 2.0, outerDiameter / 2.0 + tabHeight / 2.0 - 5.0, 0)));
                    
                    body.BooleanOperation(BooleanOperationType.BoolUnite, tab);
                    btr.AppendEntity(body);
                    tr.AddNewlyCreatedDBObject(body, true);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du diaphragme : {ex.Message}");
                    body.Dispose();
                }

                DBPoint port1 = new DBPoint(Point3d.Origin);
                DBPoint port2 = new DBPoint(new Point3d(thickness, 0, 0));
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                tr.AddNewlyCreatedDBObject(port1, true);
                btr.AppendEntity(port2);
                tr.AddNewlyCreatedDBObject(port2, true);

                AddInvisibleAttributes(btr, tr, "ORIFICE");

                tr.Commit();
            }
            return blockId;
        }
    }
}