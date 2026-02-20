using System.Collections.Generic;

namespace LogiK3D.Piping
{
    /// <summary>
    /// Base de données contenant les dimensions standards (Face-à-Face, Diamètres, Hauteurs)
    /// pour les équipements en ligne (Vannes, Filtres, Débitmètres, etc.).
    /// Basé sur les normes EN 558 / ISO 5752 pour les vannes et filtres, et EN 1092-1 pour les brides.
    /// </summary>
    public static class InlineComponentDatabase
    {
        public class ComponentDimensions
        {
            public double Length { get; set; } // Longueur Face-à-Face (Face-to-Face)
            public double FlangeDiameter { get; set; } // Diamètre extérieur du corps/bride
            public double Height { get; set; } // Hauteur de l'actionneur, du transmetteur ou de la tige
        }

        /// <summary>
        /// Retourne le diamètre extérieur de la bride selon la norme EN 1092-1.
        /// </summary>
        public static double GetFlangeOuterDiameter(int dn, int pn)
        {
            if (dn <= 50)
            {
                switch (dn)
                {
                    case 10: return 90;
                    case 15: return 95;
                    case 20: return 105;
                    case 25: return 115;
                    case 32: return 140;
                    case 40: return 150;
                    case 50: return 165;
                }
            }
            else if (dn == 65) return 185;
            else if (dn == 80) return 200;
            else if (dn == 100) return (pn <= 16) ? 220 : 235;
            else if (dn == 125) return (pn <= 16) ? 250 : 270;
            else if (dn == 150) return (pn <= 16) ? 285 : 300;
            else if (dn == 200)
            {
                if (pn == 10) return 340;
                if (pn == 16) return 340;
                if (pn == 25) return 360;
                return 375; // PN40
            }
            else if (dn == 250)
            {
                if (pn == 10) return 395;
                if (pn == 16) return 405;
                if (pn == 25) return 425;
                return 450; // PN40
            }
            else if (dn == 300)
            {
                if (pn == 10) return 445;
                if (pn == 16) return 460;
                if (pn == 25) return 485;
                return 515; // PN40
            }
            else if (dn == 350)
            {
                if (pn == 10) return 505;
                if (pn == 16) return 520;
                if (pn == 25) return 555;
                return 580; // PN40
            }
            else if (dn == 400)
            {
                if (pn == 10) return 565;
                if (pn == 16) return 580;
                if (pn == 25) return 620;
                return 660; // PN40
            }
            else if (dn == 500)
            {
                if (pn == 10) return 670;
                if (pn == 16) return 715;
                if (pn == 25) return 730;
                return 755; // PN40
            }
            else if (dn == 600)
            {
                if (pn == 10) return 780;
                if (pn == 16) return 840;
                if (pn == 25) return 845;
                return 890; // PN40
            }
            
            return dn * 1.5 + 50; // Fallback
        }

        /// <summary>
        /// Retourne le diamètre extérieur pour les composants Wafer (qui s'insèrent à l'intérieur des boulons).
        /// </summary>
        public static double GetWaferOuterDiameter(int dn, int pn)
        {
            if (dn == 15) return 51;
            if (dn == 20) return 61;
            if (dn == 25) return 71;
            if (dn == 32) return 82;
            if (dn == 40) return 92;
            if (dn == 50) return 107;
            if (dn == 65) return 127;
            if (dn == 80) return 142;
            if (dn == 100) return (pn <= 16) ? 162 : 168;
            if (dn == 125) return (pn <= 16) ? 192 : 194;
            if (dn == 150) return (pn <= 16) ? 218 : 224;
            if (dn == 200) return (pn <= 16) ? 273 : 284;
            if (dn == 250) return (pn <= 16) ? 328 : 340;
            if (dn == 300) return (pn <= 16) ? 378 : 400;
            if (dn == 350) return (pn <= 16) ? 438 : 457;
            if (dn == 400) return (pn <= 16) ? 489 : 514;
            
            return GetFlangeOuterDiameter(dn, pn) - 40; // Fallback
        }

        /// <summary>
        /// Récupère les dimensions complètes d'un composant en fonction de son type, son DN et son PN.
        /// </summary>
        public static ComponentDimensions GetDimensions(string compType, int dn, int pn)
        {
            double length = 0;
            double height = dn * 1.5; // Hauteur par défaut
            double flangeDia = GetFlangeOuterDiameter(dn, pn);

            switch (compType.ToUpper())
            {
                case "VANNE":
                case "VALVE":
                case "VALVE_BUTTERFLY":
                    // EN 558 Série 20 (Wafer)
                    var lengthsBF = new Dictionary<int, double> { {15,33}, {20,33}, {25,33}, {32,33}, {40,33}, {50,43}, {65,46}, {80,46}, {100,52}, {125,56}, {150,56}, {200,60}, {250,68}, {300,78}, {350,78}, {400,102}, {500,127}, {600,154} };
                    length = lengthsBF.ContainsKey(dn) ? lengthsBF[dn] : dn * 0.3;
                    height = dn * 1.2 + 50;
                    break;

                case "VALVE_GLOBE":
                    // EN 558 Série 1 (Brides)
                    var lengthsGL = new Dictionary<int, double> { {10,130}, {15,130}, {20,150}, {25,160}, {32,180}, {40,200}, {50,230}, {65,290}, {80,310}, {100,350}, {125,400}, {150,480}, {200,600}, {250,730}, {300,850} };
                    length = lengthsGL.ContainsKey(dn) ? lengthsGL[dn] : dn * 3;
                    height = dn * 2 + 100;
                    break;

                case "VALVE_BALL":
                    // EN 558 Série 27 (Brides)
                    var lengthsBA = new Dictionary<int, double> { {10,115}, {15,115}, {20,120}, {25,125}, {32,130}, {40,140}, {50,150}, {65,170}, {80,180}, {100,190}, {125,325}, {150,350}, {200,400} };
                    length = lengthsBA.ContainsKey(dn) ? lengthsBA[dn] : dn * 2;
                    height = dn * 1.0 + 50;
                    break;

                case "CHECK_VALVE":
                    // Wafer Check Valve
                    var lengthsCV = new Dictionary<int, double> { {15,16}, {20,19}, {25,22}, {32,28}, {40,31.5}, {50,40}, {65,46}, {80,50}, {100,60}, {125,90}, {150,106}, {200,140}, {250,145}, {300,160} };
                    length = lengthsCV.ContainsKey(dn) ? lengthsCV[dn] : dn * 0.5;
                    height = 0;
                    flangeDia = GetWaferOuterDiameter(dn, pn);
                    break;

                case "FILTRE":
                case "FILTER":
                    // EN 558 Série 1
                    var lengthsFI = new Dictionary<int, double> { {15,130}, {20,150}, {25,160}, {32,180}, {40,200}, {50,230}, {65,290}, {80,310}, {100,350}, {125,400}, {150,480}, {200,600}, {250,730}, {300,850} };
                    length = lengthsFI.ContainsKey(dn) ? lengthsFI[dn] : dn * 3;
                    height = dn * 1.5 + 50;
                    break;

                case "DEBIMETRE":
                case "FLOWMETER":
                    var lengthsFM = new Dictionary<int, double> { {15,200}, {20,200}, {25,200}, {32,200}, {40,200}, {50,200}, {65,200}, {80,200}, {100,250}, {125,250}, {150,300}, {200,350}, {250,450}, {300,500} };
                    length = lengthsFM.ContainsKey(dn) ? lengthsFM[dn] : dn * 2;
                    height = dn * 1.0 + 150;
                    break;

                case "DIAPHRAGME":
                case "DIAPHRAGM":
                    length = (dn <= 150) ? 16 : 18;
                    height = dn * 0.5 + 50;
                    flangeDia = GetWaferOuterDiameter(dn, pn);
                    break;

                default:
                    return null;
            }

            return new ComponentDimensions { Length = length, FlangeDiameter = flangeDia, Height = height };
        }
    }
}