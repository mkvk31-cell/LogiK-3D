using System;
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
        public MainPaletteControl()
        {
            InitializeComponent();
        }

        private void BtnRoutePipe_Click(object sender, RoutedEventArgs e)
        {
            // Exécuter la commande de tracé de tuyauterie
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // On utilise SendStringToExecute pour lancer une commande AutoCAD depuis l'interface WPF
                doc.SendStringToExecute("LOGIK_ROUTE_PIPE ", true, false, false);
            }
        }

        private void BtnConvertPolyline_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
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
    }
}