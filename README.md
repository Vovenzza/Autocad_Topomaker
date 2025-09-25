# AutoCAD 3D Terrain Generator (TopoBuilder)

A powerful Autodesk AutoCAD add-in designed to generate complex 3D terrain models from 2D survey point data. This tool is perfect for civil engineers, surveyors, and designers who need to quickly create accurate 3D solids and surface meshes from text-based elevation points.

### What It Does

The script provides a simple and powerful workflow directly within AutoCAD:

1.  **Select a Sample Point:** You begin by selecting a single text object that is representative of your survey points (e.g., same layer and color).
2.  **Automatic Data Collection:** The script scans the entire drawing to find all text objects that match the properties of your sample. It intelligently extracts the **XY coordinates** from the text's insertion point and the **Z elevation** from its numerical content.
3.  **Intelligent Filtering:** If multiple points exist at the same XY location, the script automatically uses the one with the highest Z-value, preventing data conflicts.
4.  **3D Model Generation:** The script uses a Delaunay triangulation algorithm to create a topologically correct network from the points. It then builds two objects:
    *   A watertight **3D Solid** representing the terrain volume.
    *   A lightweight **PolyFaceMesh** for easy surface visualization.
5.  **Text Preparation (Optional):** An included utility command, `ALTXT`, helps you prepare your drawing by aligning elevation text to the center of its corresponding marker (circle, arc, or block).

### Key Features

*   **Intuitive Workflow:** A simple "select sample and run" process right inside the AutoCAD command line.
*   **3D Solid & Mesh Output:** Creates both a solid for analysis and a mesh for visualization in a single operation.
*   **Property-Based Filtering:** Automatically processes only the text objects that match the layer and color of your sample.
*   **Robust Triangulation Engine:** Implements a Delaunay algorithm from scratch to ensure a correct and efficient surface model.
*   **Duplicate Point Handling:** Automatically resolves stacked points by selecting the highest elevation, ensuring a clean model.
*   **Helper Utility Included:** Comes with the `ALTXT` command to simplify drawing cleanup and preparation.

### How to Use

1.  **Load the Add-in:** Load the compiled `.dll` file into your AutoCAD session using the `NETLOAD` command.
2.  **(Optional) Prepare Drawing:** Use the `ALTXT` command to align your text labels to their markers for a cleaner drawing.
3.  **Run the Main Command:** Type `TOPOMODEL` in the command line.
4.  **Select Sample Text:** Follow the prompt to select one of your elevation text objects.
5.  **Done!** The script will process all matching points and create the 3D solid and visualization mesh in your model space.

### AI-Assisted Development ("Vibecoding")

This script was developed with significant assistance from AI. The process, which could be described as **"Vibecoding,"** involved translating a clear functional vision and complex algorithmic logic into robust C# code through collaboration with an AI partner. The AI helped structure the code, implement the Delaunay triangulation and solid geometry functions, and debug the workflow, acting as a powerful tool to bring the initial idea to life.

### Installation & Setup

To use this script, you need to compile it and load it into AutoCAD.

1.  **Compile the Code:** Open the project in Visual Studio (2019 or later) and build the solution. This will create a `TopoBuilder.dll` file in the `bin/Debug` or `bin/Release` folder.
2.  **Load into AutoCAD:**
    *   Open AutoCAD.
    *   Type `NETLOAD` in the command line.
    *   A file browser window will open. Navigate to the location of your `TopoBuilder.dll` file (e.g., in the `bin/Debug` folder) and select it.
    *   The add-in is now loaded for your current session. You will need to `NETLOAD` it again if you restart AutoCAD.

---

**Disclaimer:** This is a utility script. Always back up your drawings before running commands that create or modify a large amount of geometry. Use at your own risk.
