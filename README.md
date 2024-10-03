# SolidWorks Spline from Points

## Overview
This script creates a spline in a 3D sketch, using coordinates from a CSV file as control points in SolidWorks.

## Requirements
- SolidWorks with an open part document.
- CSV file containing coordinates in the format: `"Bone Name", "X", "Y", "Z"`.

## Usage
1. Open the Visual Basic for Applications (VBA) editor in SolidWorks (`Tools > Macro > New`).
2. Copy and paste the code provided in `SW_spline_points.vb` into the new macro.
3. Update the `filePath` variable in the script to point to your CSV file containing the coordinates.
4. Save the macro as an `.swp` file.
5. Run the macro (`Tools > Macro > Run`) in SolidWorks while the part document is active.
6. The script will create a 3D sketch and generate a spline based on the points from the CSV file.

## Notes
- Ensure that the CSV file contains enough points (at least two) to create a spline.
- You must update the `filePath` in the code to specify the correct location of your CSV file.
