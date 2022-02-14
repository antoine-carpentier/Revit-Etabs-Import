# Revit-Etabs-Import
Reads CSV files created by [ETABSModelTransfer](https://github.com/antoine-carpentier/ETABSModelTransfer) to create a Revit model.

## How to Use

1. Download this repository's files
2. Open up your chosen Revit file
3. Create a new Module in Manage > Macro Manager
4. When SharpDevelop opens up, right-click your project and select "Add Existing Items".
5. Select `ThisDocument.vb`, `clsCollectors.vb`, `frmModelTransfer.Designer.vb` and `frmModelTransfer.vb`.
6. Add References to `LumenWorks.Framework.IO.dll` from this repository, as well as `System.Drawings` and `System.Windows.Forms`.
7. Build.

## Video

Watch [Youtube Video](https://www.youtube.com/watch?v=RYmr9pq0Kio) to see [ETABSModelTransfer](https://github.com/antoine-carpentier/ETABSModelTransfer) and [Revit-Etabs-Import](https://github.com/antoine-carpentier/Revit-Etabs-Import) working in tandem.
