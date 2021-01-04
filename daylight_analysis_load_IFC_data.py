"""
This is part 1 of the daylight program. It gets and prints info about IFC model 
to excel, but can also format it, so it can be used for analysis (part 2).

Outputs
-------
Excel file
"""

# Main dependencies are xlsxwriter, ifcopenshell, numpy and matplotlib
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import xlsxwriter
import ifcopenshell
import ifcopenshell.util
from ifcopenshell.util.selector import Selector
import ifcopenshell.geom
from dataclasses import dataclass, InitVar


# IFC model used throughout the program
model = ifcopenshell.open("model/Duplex_A_20110907_optimized.ifc")


# IFC open shell geometry tree and selector lib settings
tree_settings = ifcopenshell.geom.settings()
tree_settings.set(tree_settings.DISABLE_OPENING_SUBTRACTIONS, True)
t = ifcopenshell.geom.tree()
t.add_file(model, tree_settings)

selector = Selector()

# define spaces
spaces = model.by_type("IfcSpace")


@dataclass
class SpaceParams:
    """
    Object for formatting space attributes for analysis.
    Does nothring when excelformat is chosen.
    """

    out: list
    excelFormat: InitVar[str] = None

    def __post_init__(self, excelFormat):

        if excelFormat == False:
            analysisOut = [
                val[1:2] for val in (self.out if self.out is not None else [])
            ]
            self.out = [val for sublist in analysisOut for val in sublist]
        else:
            pass
        return self.out


#############
### .XLSX ###
#############

# Set xlxswriter workbook and formats
workbook = xlsxwriter.Workbook("output/output_IFC_data.xlsx")
workbook.set_size(3000, 1500)
cell_format = workbook.add_format({"bold": True, "bg_color": "#d3d3d3"})
cell_format2 = workbook.add_format({"bold": True, "bg_color": "#7e7e7e"})


def getMaterialAndQuantities(ifcElement):
    """
    Gets material layer name and thickness for IFC elements, that are not roof.
    Here used only on walls.

    Parameters
    ----------
    ifcElement: a single instance of an IFC element, eg. IFC wall

    Returns
    -------
    An array of arrays formatted for writing to excel with column number always at index n,0.
    First array gives total number of materials in element.
    The following give alternately material name and material thickness.
    """
    out = []
    col = 0
    most_materials = 0

    if not "Roof" in str(ifcElement.Name):
        col = 5
        for relAssociatesMaterial in ifcElement.HasAssociations:
            # get the name and thickness for each materiallayer
            num_materials = len(
                relAssociatesMaterial.RelatingMaterial.ForLayerSet.MaterialLayers
            )
            if num_materials > most_materials:
                most_materials = num_materials
            out.append([col, num_materials])
            col += 1
            for (
                MaterialLayers
            ) in relAssociatesMaterial.RelatingMaterial.ForLayerSet.MaterialLayers:
                out.append([col, str(MaterialLayers.Material.Name)])
                out.append([col + 1, str(MaterialLayers.LayerThickness)])
                col += 2

    return out


def wallParams(wall, debug=False):
    """
    Get and formats information about walls, including name, tag, external or not.
    Calls material quantities function.

    Parameters
    ----------
    ifcElement: a single instance of an IFC wall

    Returns
    -------
    An array of arrays formatted for writing to excel with column number always at index n,0.
    """

    outWall = []
    outWall.append([2, wall.Name])
    outWall.append([3, wall.Tag])

    for relDefinesByProperties in wall.IsDefinedBy:
        if "IfcRelDefinesByProperties" in str(relDefinesByProperties):
            wallProps = relDefinesByProperties.RelatingPropertyDefinition.HasProperties
            for prop in wallProps:
                if prop.Name == "IsExternal":
                    outWall.append([4, prop.NominalValue.wrappedValue])

    for i in getMaterialAndQuantities(wall):
        outWall.append(i)

    return outWall


def windowParams(window, debug=False, excelFormat=False):
    """
    Get and formats information about windows for printing to excel and for simulation.

    Parameters
    ----------
    ifcElement: a single instance of an IFC window
    excelformat: True or False

    Returns
    -------
    When excelformat is true: An array of arrays formatted for writing to excel
    with column number always at index n,0.

    When excelformat is false: Appart from the standard information, the name and
    length is also added. So are the point of location of window on wall.

    """
    outWindow = []
    windowHeight = selector.get_element_value(
        window, "PSet_Revit_Type_Dimensions.Height"
    )
    windowWidth = selector.get_element_value(window, "PSet_Revit_Type_Dimensions.Width")
    sillHeight = selector.get_element_value(
        window, "PSet_Revit_Constraints.Sill Height"
    )
    windowLocation = window.ObjectPlacement.PlacementRelTo.RelativePlacement.Location[0]
    loc_x, loc_y = windowLocation[0], windowLocation[2]

    wall_name = ""
    wall_length = 0
    for wall in t.select_box(window):
        if wall.is_a("IfcWallStandardCase"):

            wall_length = selector.get_element_value(
                wall, "PSet_Revit_Dimensions.Length"
            )
            # find wall orientation, front is (assumed) North
            placement = wall.ObjectPlacement.RelativePlacement.RefDirection

            if placement is None:
                placement = ((0.0, 0.0, 0.0),)
            placement = placement[0]
            if placement == (0.0, 0.0, 0.0):
                wall_name = "front"
            elif placement == (0.0, -1.0, 0.0):
                wall_name = "right"
            elif placement == (0.0, 1.0, 0.0):
                wall_name = "left"
            elif placement == (-1.0, 0.0, 0.0):
                wall_name = "back"

    if excelFormat == True:
        outWindow.append([2, window.Name])
        outWindow.append([3, window.Tag])
        outWindow.append([4, round(windowHeight, 3)])
        outWindow.append([5, round(windowWidth, 3)])
        outWindow.append([6, sillHeight])
    else:
        outWindow.append(window.Tag)
        outWindow.append(round(windowHeight, 3))
        outWindow.append(round(windowWidth, 3))
        try:
            outWindow.append(round(sillHeight, 3))
        except:
            outWindow.append(sillHeight)
        outWindow.append(wall_name)
        outWindow.append(wall_length)
        outWindow.append(loc_x)
        outWindow.append(loc_y)

    return outWindow


def doorParams(door):
    """
    Get and formats information about doors, including name, tag, glass or not.

    Parameters
    ----------
    ifcElement: a single instance of an IFC door

    Returns
    -------
    An array of arrays formatted for writing to excel with column number always at index n,0.
    """
    outDoor = []
    outDoor.append([2, door.Name])
    outDoor.append([3, door.Tag])

    if "Glass" in door.Name:
        doorHeight = door.OverallHeight
        doorWidth = door.OverallWidth

        outDoor.append([4, "External Glass Door"])
        outDoor.append([5, round(doorHeight, 3)])
        outDoor.append([6, round(doorWidth, 3)])

    else:
        doorHeight = door.OverallHeight
        doorWidth = door.OverallWidth

        outDoor.append([4, "External No-glass Door"])
        outDoor.append([5, round(doorHeight, 3)])
        outDoor.append([6, round(doorWidth, 3)])
    return outDoor


def intersectingObjects(
    space,
    debug=False,
    windowsOnly=False,
    doorsOnly=False,
    wallsOnly=False,
    excelFormat=False,
):
    """
    Get and formats information about spaces.
    Finds windows, doors and walls bounding a space.
    Calls the respective IFC element functions.

    Parameters
    ----------
    ifcElement: a single instance of an IFC space.
    windowsOnly: appends only info about spaces and windows
    doorsOnly: appends only info about spaces and doors
    wallsOnly: appends only info about spaces and walls

    Returns
    -------
    An array of arrays formatted for writing to excel with column number always at index n,0.
    """
    out = []
    if excelFormat == True:
        out.append(
            [
                [0, space.LongName, cell_format],
                [1, space.Name, cell_format],
                [2, " ", cell_format],
                [3, " ", cell_format],
                [4, " ", cell_format],
                [5, " ", cell_format],
                [6, " ", cell_format],
            ]
        )
    else:
        out.append(space.LongName)
        out.append(space.Name)

    # get windows and doors that intersect space bounding box (space.BoundedBy doesn't get all windows)
    for obj in t.select_box(space, extend=0.5):
        if obj.is_a("IfcWindow") and windowsOnly:
            out.append(windowParams(obj, debug, excelFormat))
        elif obj.is_a("IfcDoor") and doorsOnly:
            for relDefinesByProperties in obj.IsDefinedBy:

                if "IfcRelDefinesByProperties" in str(relDefinesByProperties):
                    doorProps = (
                        relDefinesByProperties.RelatingPropertyDefinition.HasProperties
                    )
                    # get external doors only
                    for prop in doorProps:
                        if prop.Name == "IsExternal":
                            if prop.NominalValue.wrappedValue == True:
                                out.append(doorParams(obj))

    # get all walls that bound a space
    # print("\n\t####{}\n".format(space.Name))
    for obj in space.BoundedBy:
        if obj.RelatedBuildingElement != None and wallsOnly:
            if obj.RelatedBuildingElement.is_a("IfcWall"):
                out.append(wallParams(obj.RelatedBuildingElement))
    return out


def arbiClosOut(prop, points):
    """
    Gets the points of a non rectangular space profile,
    creates a bounding box and gets the dimensions of it.

    Parameters
    ----------
    prop: space property "IfcArbitraryClosedProfileDef"
    points: empty list

    Returns
    -------
    An array of arrays formatted for writing to excel with column number always at index n,0.
    """
    for point in prop.OuterCurve.Points:
        points.append(point[0])

    bot_left_x = min(point[0] for point in points)
    bot_left_y = min(point[1] for point in points)
    top_right_x = max(point[0] for point in points)
    top_right_y = max(point[1] for point in points)

    XDim = top_right_x - bot_left_x
    YDim = top_right_y - bot_left_y
    return [[2, round(XDim, 3), cell_format], [3, round(YDim, 3), cell_format]]


def spaceDims(space):
    """
    Gets the name, code, x, y and z dimension of space.
    Different methods for spaces with rectangular profile and irregular.
    Hallway and Roof are excluded.

    Parameters
    ----------
    space: a single instance of an IFC space.

    Returns
    -------
    An array of arrays formatted for writing to excel with column number always at index n,0.
    """
    points = []
    if space.LongName != "Hallway" and space.LongName != "Roof":

        prop = space.Representation.Representations[0][3][0][0]
        out = []
        out.append([0, space.LongName])
        out.append([1, space.Name])

        if "IfcRectangleProfileDef" in str(prop):
            a, b = [[2, round(prop.XDim, 3)], [3, round(prop.YDim, 3)]]
            out.append(b)
            out.append(a)

        elif "IfcArbitraryClosedProfileDef" in str(prop):
            a, b = arbiClosOut(prop, points)
            out.append(a)
            out.append(b)

        spaceHeight = selector.get_element_value(
            space, "PSet_Revit_Dimensions.Unbounded Height"
        )
        out.append([4, round(spaceHeight, 3)])

        return out
    else:
        pass


def excelWrite(sheet, lines, debug=False):
    """
    Gets the formatted arrays for printing to excel.
    Adjusts the row number if formatting is added to a cell.

    Parameters
    ----------
    sheet: Sheet is a object referencing the excel worksheet
    lines: Lines is a list of queries to print

    Returns
    -------
    Prints to excel
    """

    i_adjust = 0
    for i in range(0, len(lines)):

        if lines[i] == None:
            i_adjust = i_adjust - 1

        else:
            for val in lines[i]:

                if debug:
                    print("val:", val)
                if len(val) == 2:
                    col, text = val
                    if debug:
                        print(col, text)
                    sheet.write(i + 1 + i_adjust, col, text)
                elif len(val) == 3:
                    col, text, formatting = val
                    sheet.write(i + 1 + i_adjust, col, text, formatting)
                else:
                    raise Exception(
                        "DataError",
                        "excelWrites second argument should have length 2 or 3!",
                    )


def spaceFunc(
    spaces,
    functions,
    nested=False,
    windowsOnly=False,
    doorsOnly=False,
    wallsOnly=False,
    excelFormat=False,
):
    """
    Calls the chosen functions and formats them correctly for printing to excel or for analysis for each space.

    Parameters
    ----------
    spaces: IFC spaces
    functions: Function to be run
    nested=False, #TODO add decription
    windowsOnly=False,
    doorsOnly=False,
    wallsOnly=False,
    excelFormat=False,

    Returns
    -------
    List to be used by excelWrite
    """

    out = []

    for space in spaces:
        for f in functions:
            if nested:
                if excelFormat:
                    rows = f(
                        space,
                        windowsOnly=windowsOnly,
                        doorsOnly=doorsOnly,
                        wallsOnly=wallsOnly,
                        excelFormat=excelFormat,
                    )
                    for val in rows:
                        out.append(val)
                else:
                    rows = f(
                        space,
                        windowsOnly=windowsOnly,
                        doorsOnly=doorsOnly,
                        wallsOnly=wallsOnly,
                        excelFormat=excelFormat,
                    )
                    out.append(rows)
            else:
                out.append(f(space))
    return out


def getMaterialAndQuantitiesHeaders(spaces):
    """
    Find the number of materials in wall material layer.

    Parameters
    ----------
    spaces: IFC spaces

    Returns
    -------
    Number of materials

    """
    most_materials = 0
    for space in spaces:
        for obj in space.BoundedBy:
            if obj.RelatedBuildingElement != None:
                if obj.RelatedBuildingElement.is_a("IfcWall"):
                    if not "Roof" in str(obj.RelatedBuildingElement.Name):
                        for (
                            relAssociatesMaterial
                        ) in obj.RelatedBuildingElement.HasAssociations:
                            num_materials = len(
                                relAssociatesMaterial.RelatingMaterial.ForLayerSet.MaterialLayers
                            )
                            if num_materials > most_materials:
                                most_materials = num_materials
    return most_materials


############
### MAIN ###
############


def main(debug=False):
    """
    Main functions that builds excel sheets and calls the correct write functions.

    Parameters
    ----------

    Returns
    -------

    """
    spaces = model.by_type("IfcSpace")

    # ASSUMPTIONS SHEET
    assumptionsSheet = workbook.add_worksheet("Assumptions")
    assumptionsSheet.set_column(0, 4, 15)
    assumptionsSheet.write(0, 0, "Assumptions", cell_format2)
    assumptionsSheet.write(1, 0, "Thermal conductivity", cell_format)
    assumptionsSheet.write(2, 1, "Windows")
    assumptionsSheet.write(3, 1, "Glass Doors")
    assumptionsSheet.write(4, 1, "Non-glass External Doors")
    assumptionsSheet.write(5, 1, "External Walls")
    assumptionsSheet.write(2, 2, "1,2 W/m²K")
    assumptionsSheet.write(3, 2, "1,5 W/m²K")
    assumptionsSheet.write(4, 2, "1,4 W/m²K")
    assumptionsSheet.write(5, 2, "0,09 W/m²K")
    assumptionsSheet.write(
        6,
        0,
        "Spaces with non-rectangular floor profile will be analyzed based on their bounding box",
        cell_format,
    )
    assumptionsSheet.write(
        7,
        0,
        "Example of Foyer floor profile and the bounding box points that will be used to create a new energy zone",
    )
    assumptionsSheet.insert_image(
        "B12", "output/foyerpoints.png", {"x_offset": 15, "y_offset": 10}
    )

    # SPACES SHEET
    spaceDimSheet = workbook.add_worksheet("Spaces")
    spaceDimSheet.set_column(0, 4, 15)
    spaceDimSheet.write(0, 0, "Space Name", cell_format2)
    spaceDimSheet.write(0, 1, "Space Code", cell_format2)
    spaceDimSheet.write(0, 2, "X Dimension", cell_format2)
    spaceDimSheet.write(0, 3, "Y Dimension", cell_format2)
    spaceDimSheet.write(0, 4, "Height", cell_format2)
    spaceDimSheet.write(25, 2, "", cell_format)
    spaceDimSheet.write(25, 3, "Based on bounding box")

    excelWrite(spaceDimSheet, SpaceParams(spaceFunc(spaces, [spaceDims])).out, debug)

    # WINDOW SHEET
    windowSheet = workbook.add_worksheet("Windows")

    windowSheet.set_column(0, 4, 15)
    windowSheet.set_column(2, 2, 45)
    windowSheet.write(0, 0, "Space Name", cell_format2)
    windowSheet.write(0, 1, "Space Code", cell_format2)
    windowSheet.write(0, 2, "Winow Name", cell_format2)
    windowSheet.write(0, 3, "Window Tag", cell_format2)
    windowSheet.write(0, 4, "Height", cell_format2)
    windowSheet.write(0, 5, "Width", cell_format2)
    windowSheet.write(0, 6, "Sill Height", cell_format2)

    excelWrite(
        windowSheet,
        spaceFunc(
            spaces,
            [intersectingObjects],
            nested=True,
            windowsOnly=True,
            excelFormat=True,
        ),
    )

    # DOOR SHEET
    doorSheet = workbook.add_worksheet("External Doors")

    doorSheet.set_column(0, 4, 15)
    doorSheet.set_column(2, 2, 45)
    doorSheet.write(0, 0, "Space Name", cell_format2)
    doorSheet.write(0, 1, "Space Code", cell_format2)
    doorSheet.write(0, 2, "External Door Name", cell_format2)
    doorSheet.write(0, 3, "Door Tag", cell_format2)
    doorSheet.write(0, 4, "Type", cell_format2)
    doorSheet.write(0, 5, "Height", cell_format2)
    doorSheet.write(0, 6, "Width", cell_format2)

    excelWrite(
        doorSheet,
        spaceFunc(
            spaces, [intersectingObjects], nested=True, doorsOnly=True, excelFormat=True
        ),
        debug,
    )

    # WALL SHEET
    wallSheet = workbook.add_worksheet("Walls")

    wallSheet.set_column(0, 4, 15)
    wallSheet.set_column(2, 2, 50)
    wallSheet.write(0, 0, "Space Name", cell_format2)
    wallSheet.write(0, 1, "Space Code", cell_format2)
    wallSheet.write(0, 2, "Wall Name", cell_format2)
    wallSheet.write(0, 3, "Wall Tag", cell_format2)
    wallSheet.write(0, 4, "Is external?", cell_format2)
    wallSheet.write(0, 5, "# layers", cell_format2)
    material = 1
    most_materials = getMaterialAndQuantitiesHeaders(spaces)

    for a in range(6, (most_materials * 2) + 6, 2):
        wallSheet.write(0, a, "Material {}".format(material), cell_format2)
        wallSheet.write(0, a + 1, "Thickness", cell_format2)
        wallSheet.set_column(a, a, 25)
        material += 1

    excelWrite(
        wallSheet,
        spaceFunc(
            spaces, [intersectingObjects], nested=True, wallsOnly=True, excelFormat=True
        ),
        debug,
    )

    return


print(main(debug=False))

workbook.close()
