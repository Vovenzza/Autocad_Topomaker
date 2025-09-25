AutoCAD 3D Terrain Generator (TopoBuilder)

This is a plugin for AutoCAD designed to create detailed 3D terrain models from topographical survey data. The script processes text objects representing survey points to generate a Solid3d model and a visual PolyFaceMesh.

The project also includes a handy utility command to align text labels with their corresponding survey markers, making it a complete solution for working with survey data in AutoCAD.
Features

    3D Solid Terrain Generation: Creates a watertight Solid3d object from point data, perfect for volumetric analysis, sectioning, or further 3D modeling.

    Surface Visualization Mesh: Generates a lightweight PolyFaceMesh on a separate layer, providing a quick and clear visualization of the terrain's surface.

    Data-Driven Modeling: Intelligently builds the model by:

        Using the XY insertion point of text objects for location.

        Using the numerical content of the text for Z elevation.

    Intelligent Point Filtering:

        Allows the user to select a sample text object, and the script will only process other text entities that match the sample's layer and color.

        Automatically resolves points with duplicate XY coordinates by selecting the one with the highest Z elevation, ensuring a clean surface.

    Delaunay Triangulation: Implements a robust algorithm to create a topologically correct Triangulated Irregular Network (TIN) from the point data, forming the basis of the 3D model.

    Auxiliary Text Alignment Tool: Includes a helper command to quickly center text labels on their corresponding survey markers (circles, arcs, or blocks), simplifying drawing cleanup before generating the model.

Available Commands
TOPOMODEL (Main Command)

This is the primary command of the plugin. It initiates the entire process of creating the 3D terrain solid and the accompanying surface mesh.
ALTXT (Utility Command)

This is a helper tool used to prepare a drawing. It aligns text objects to the center of a sample circle, arc, or block, ensuring that elevation labels are correctly positioned over their markers.
Workflow & How to Use

    Prepare Your Drawing:

        Ensure your drawing contains text objects where the XY position represents the survey point and the text's content is the elevation value (e.g., "123.45").

        All survey point text objects should share common properties, like being on the same layer.

        (Optional) If your text labels are not centered on their markers, use the ALTXT command to clean up their positions first.

    Load the Plugin:

        Open AutoCAD.

        Type the NETLOAD command into the command line.

        Browse to and select the compiled TopoBuilder.dll file.

    Run the Terrain Generator:

        Type TOPOMODEL in the command line and press Enter.

        The command prompt will ask you to "Select sample text object...". Click on one of the text objects that you want to use for the terrain generation.

        The script will automatically read its properties (layer, color) and use them to find all other matching survey points.

    Review the Output:

        The script will process all valid points and generate two new objects:

            A 3D Solid representing the terrain volume.

            A PolyFaceMesh for visualization (typically on a new "Topo_Surface_Visualization" layer).

Disclaimer

This script was written with the assistance of an AI programming partner (vibecoding).