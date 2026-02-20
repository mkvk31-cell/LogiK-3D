# LogiK 3D

LogiK 3D est un plugin AutoCAD .NET (compatible AutoCAD 2025/2026) conçu pour faciliter le routage de tuyauterie 3D et la génération de fichiers isométriques (PCF), sans nécessiter l'environnement de projet lourd de Plant 3D.

## Fonctionnalités Principales

*   **Palette WPF Interactive** : Une interface utilisateur complète intégrée à AutoCAD pour sélectionner les diamètres (ISO/DIN) et insérer des composants.
*   **Routage 3D Personnalisé** : Dessinez des tuyaux en 3D (Solid3d) avec génération automatique des coudes, sans utiliser les objets complexes de Plant 3D.
*   **Conversion de Polylignes** : Transformez instantanément des polylignes 3D en réseaux de tuyauterie solides.
*   **Export PCF Indépendant** : Générez des fichiers `.pcf` standards pour la création d'isométriques, en lisant les données étendues (XData) attachées aux solides AutoCAD.
*   **Lecture de Spécifications Plant 3D** : Lisez directement les fichiers de spécification `.pspx` de Plant 3D via l'API `PnP3dPartsRepository` pour récupérer les diamètres extérieurs exacts, sans avoir besoin d'une base de données SQLite ou d'un projet Plant 3D actif.

## Prérequis

*   AutoCAD 2025 ou 2026
*   AutoCAD Plant 3D 2026 (pour l'accès aux bibliothèques de spécifications et aux DLLs)
*   .NET 8.0 SDK

## Installation et Utilisation

1.  Compilez le projet avec Visual Studio ou via la ligne de commande (`dotnet build`).
2.  Dans AutoCAD, utilisez la commande `NETLOAD` et sélectionnez le fichier `LogiK3D.dll` généré dans le dossier `bin/Debug/net8.0-windows/`.
3.  Tapez la commande `LOGIK_PALETTE` pour afficher l'interface utilisateur.
4.  Utilisez les boutons de la palette pour router des tuyaux ou exporter un fichier PCF.

## Architecture Technique

*   **Framework** : .NET 8.0 (C#)
*   **UI** : WPF (Windows Presentation Foundation) hébergé dans un `PaletteSet` AutoCAD.
*   **Données** : Utilisation des `XData` d'AutoCAD pour stocker les informations métier (Code SAP, Longueur, DN) directement sur les entités `Solid3d`.
*   **Intégration Plant 3D** : Utilisation exclusive de `Autodesk.ProcessPower.PartsRepository.Specification.PipePartSpecification` pour lire les catalogues, évitant ainsi les contraintes de `PnPProjectManager`.
