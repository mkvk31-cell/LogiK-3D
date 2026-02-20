using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using LogiK3D.Piping;

namespace LogiK3D.UI
{
    public partial class MainPaletteControl : UserControl
    {
        // Propriétés statiques pour partager les données avec les commandes
        public static double CurrentOuterDiameter { get; private set; } = 114.3;
        public static string CurrentDN { get; private set; } = "DN100";

        public MainPaletteControl()
        {
            InitializeComponent();
            LoadDiameters();
        }

        private void LoadDiameters()
        {
            // Liste des diamètres ISO/DIN (EN 10220 / ISO 4200 / DIN 2448)
            var diameters = new Dictionary<string, double>
            {
                {"DN10 (17.2 mm)", 17.2},
                {"DN15 (21.3 mm)", 21.3},
                {"DN20 (26.9 mm)", 26.9},
                {"DN25 (33.7 mm)", 33.7},
                {"DN32 (42.4 mm)", 42.4},
                {"DN40 (48.3 mm)", 48.3},
                {"DN50 (60.3 mm)", 60.3},
                {"DN65 (76.1 mm)", 76.1},
                {"DN80 (88.9 mm)", 88.9},
                {"DN100 (114.3 mm)", 114.3},
                {"DN125 (139.7 mm)", 139.7},
                {"DN150 (168.3 mm)", 168.3},
                {"DN200 (219.1 mm)", 219.1},
                {"DN250 (273.0 mm)", 273.0},
                {"DN300 (323.9 mm)", 323.9},
                {"DN350 (355.6 mm)", 355.6},
                {"DN400 (406.4 mm)", 406.4},
                {"DN450 (457.0 mm)", 457.0},
                {"DN500 (508.0 mm)", 508.0},
                {"DN600 (610.0 mm)", 610.0}
            };

            // Exemple d'utilisation de SpecReader (à adapter avec le chemin réel de la spec)
            // string specPath = @"C:\AutoCAD Plant 3D 2026 Content\CPak DIN\DIN.pspx";
            // double od = SpecReader.GetPipeOuterDiameter(specPath, "100");
            // if (od > 0) { /* utiliser od */ }

            foreach (var kvp in diameters)
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Content = kvp.Key;
                item.Tag = kvp.Value;
                CmbDiameter.Items.Add(item);
                
                // Sélectionner DN100 par défaut
                if (kvp.Key.StartsWith("DN100"))
                {
                    CmbDiameter.SelectedItem = item;
                }
            }
        }

        private void CmbDiameter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbDiameter.SelectedItem is ComboBoxItem item)
            {
                CurrentOuterDiameter = (double)item.Tag;
                CurrentDN = item.Content.ToString().Split(' ')[0];
            }
        }

        private void BtnRoutePipe_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // On utilise notre propre commande de routage interactif
                // qui ne nécessite pas d'être dans un projet Plant 3D
                doc.SendStringToExecute("LOGIK_ROUTE_PIPE ", true, false, false);
            }
        }

        private void BtnConvertPolyline_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // On utilise notre propre commande de conversion de polyligne
                // qui ne nécessite pas d'être dans un projet Plant 3D
                doc.SendStringToExecute("LOGIK_PIPE ", true, false, false);
            }
        }

        private void BtnInsertComponent_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null && btn.Tag != null)
            {
                string componentType = btn.Tag.ToString();
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.SendStringToExecute($"LOGIK_INSERT_COMP {componentType} ", true, false, false);
                }
            }
        }

        private void BtnGetInfo_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.SendStringToExecute("LOGIK_GET_INFO ", true, false, false);
            }
        }

        private void BtnExportPCF_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // Lance la commande d'export PCF
                doc.SendStringToExecute("LOGIK_EXPORT_PCF ", true, false, false);
            }
        }
    }
}