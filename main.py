"""
This is part 2 of the daylight program. It imports and formats info
from part 1, creates the correct classes, runs the analysis, post 
processes the results and has a simple terminal user interface.

Outputs
-------
Results and daylight plot if chosen.
"""

# Main dependencies are honeybee plus lib, radiance and ifcopenshell
import ifcopenshell
from honeybee_plus.room import Room
from honeybee_plus.radiance.material.glass import Glass
from honeybee_plus.radiance.properties import RadianceProperties
from honeybee_plus.radiance.sky.certainIlluminance import CertainIlluminanceLevel
from honeybee_plus.radiance.recipe.pointintime.gridbased import GridBased
from honeybee_radiance_folder import ModelFolder
from honeybee_plus.hbsurface import HBSurface
from honeybee_plus.hbfensurface import HBFenSurface
from daylight_analysis_load_IFC_data import *
import os
from dataclasses import dataclass, InitVar


# Set radiance result folder to "room"
rf = "./room"
folder = ModelFolder(rf)
folder.write(overwrite=True)


#########################################
###CREATE WINDOWOUT AND SPACEOUT LISTS###
#########################################

# Temporary variable for storing spaces
out = []

# remove excel formatting for each space.
for i in spaces:
    out.append(SpaceParams((spaceDims(i)), excelFormat=False).out)

# remove empty space entries
spaceOut = [x for x in out if x != []]

# finds windows bounding space and formats for analysis
windowOut = SpaceParams(
    spaceFunc(
        spaces,
        [intersectingObjects],
        nested=True,
        windowsOnly=True,
        excelFormat=False,
    )
).out

# remove empty window entries in window param list (windowOut)
windowOut = [y for y in (x for x in windowOut if x != [])]


@dataclass
class Window:
    """
    A window object with IFC window parameters used by the 'Spaces' class.
    Attached to a Space

    Parameters
    ----------
    name: a string containing the window ifc name
    wx: (float) window x-dimension width
    wy: (float)  window y-dimension height
    sillHeight: (float)  sill height (height from floor to bottom of window)
    wall_name: (str) parent wall orientation as string, one of ('left','right','front','back'), 'front' is North
    wall_length: (float) parent wall length
    loc_x: (float)  bottom left location point of window - x value (width)
    loc_y: (float) bottom left location point of window - y value (height)

    Returns
    -------
    Does not return anything
    """

    name: str
    wx: float
    wy: float
    sillHeight: float
    wall_name: str
    wall_length: float
    loc_x: float
    loc_y: float


@dataclass
class Space:
    """
    A space object with IFC space parameters used by the 'Spaces' class.
    Has Window attached to it.

    Parameters
    ----------
    longName: (str) Descriptive name of space. Eg. "Bedroom".
    name: (str)  Unique space code Eg. "A203"
    windows: (list) List of all window objects belonging to a given space
    sx: (float) Space width
    sy: (float) Space depth
    sz: (float) Space Height

    Returns
    -------
    Does not return anything
    """

    longName: str
    name: str
    windows: list
    sx: float
    sy: float
    sz: float


class Analysis:
    """
    Main analysis object. Creates honeybee room.
    Adds windows (+ window params) to the correct walls based on wall name and location points.
    Runs overcast sky analysis using Radiance.

    Parameters
    ----------
    spaceOut: list of spaces and params used by Spaces object
    windowOut: list of windows and params used by Spaces object
    spacename: (str) name of the space for analysis
    lightTr: (float) visible light transmittance value
    gridsize: (float) gridsize value
    analysisPlaneHeight: (float) analysis plane height value

    Returns
    -------
    results

    """

    def __createRoom(self, space):
        """
        Create a honeybee room with the same dimensions as IFC space,
        with origin in (0,0,0) and no rotation angle, since the Duplex model
        is assumed to have y-axis = North

        Parameters
        ----------
        space:

        Returns
        -------
        results
        """
        room = Room(
            origin=(0, 0, 0),
            width=space.sx,
            depth=space.sy,
            height=space.sz,
            rotation_angle=0,
        )

        return room

    def __addWindows(self, space, window):
        """ Add window to room inside a space """
        # Define Radiance glass material with light transmitance defined by user
        glass_type = Glass.by_single_trans_value("transValue", self.lightTr)
        radprops = RadianceProperties(material=glass_type)

        # From surfaces created by __createRoom function, find the wall
        # that window lies on (eg. 'back') and set it to hbsurface
        for i in space.room.surfaces:
            if window.wall_name in str(i):
                hbsurface = i

        # !COMPLICATED! - since IFC wall lenght ≠ space wall lenght
        # and window location point can be measured from either side (arbitrary), if the window location point
        # falls outside (is larger than) either space wall length, take the IFC wall lenght - (minus)
        # the location point - (minus) the window width.
        if window.loc_x > space.sx and window.loc_x > space.sy:
            window.loc_x = window.wall_length - window.loc_x - window.wx

        # Construct glazing points:
        # because HBFenSurface needs glazing points defined by (x, y, z) for each point,
        # where y is the 'depth' of wall axis, it is set to 0 here.
        # The remaining three points are constructed from the original bottom left point, by adding window length/height.
        glzpts = [
            (window.loc_x, 0, window.loc_y),
            (window.loc_x + window.wx, 0, window.loc_y),
            (
                window.loc_x + window.wx,
                0,
                window.loc_y + window.wy,
            ),
            (window.loc_x, 0, window.loc_y + window.wy),
        ]
        # Construct glazing surface from glazing points with same name as window
        glzsrf = HBFenSurface(str(window.name), glzpts, rad_properties=radprops)
        # Add glazing surface to the honeybee wall surface
        hbsurface.add_fenestration_surface(glzsrf)

    def __init__(
        self,
        spaceOut,
        windowOut,
        spacename,
        lightTr=0.6,
        gridsize=0.5,
        analysisPlaneHeight=0.75,
    ):
        """Initialise Analysis. Create a honeybee room for each space and add windows to each wall of space"""
        self.Spaces = Spaces(spaceOut, windowOut)
        self.spacename = spacename
        self.lightTr = lightTr
        self.gridsize = gridsize
        self.analysisPlaneHeight = analysisPlaneHeight
        for space in self.Spaces.spaces:
            space.room = self.__createRoom(space)
            for window in space.windows:
                self.__addWindows(space, window)

    def returnAnalysis(self):
        """Run the gridbased daylight simulation using Radiance for space chosen by user."""

        for space in self.Spaces.spaces:
            if space.name == self.spacename:

                ####################################################################################
                ### This part is based on example from https://github.com/ladybug-tools/honeybee ###
                ####################################################################################

                # run a grid-based analysis for this room
                # generate an overcast sky with 10 000 lux
                sky = CertainIlluminanceLevel(illuminance_value=10000)

                # generate grid of test points with grid size and analysis plane height chosen by user
                analysis_grid = space.room.generate_test_points(
                    grid_size=self.gridsize,
                    height=self.analysisPlaneHeight,
                )

                # put the recipe together
                rp = GridBased(
                    sky=sky,
                    analysis_grids=(analysis_grid,),
                    simulation_type=0,
                    hb_objects=(space.room,),
                )

                # write simulation to folder
                batch_file = rp.write(target_folder=".", project_name="room")

                # run the simulation
                rp.run(batch_file, debug=True)

                # results - in this case it will be an analysis grid
                result = rp.results()[0]
        return result


class Spaces:
    """
    Space object: maps the correct parameters into the Space and Window classes.
    Can print the objects and their parameters.

    Parameters
    ----------
    spaceOut: (list) list of spaces and their params
    windowOut: (list) list of windows and their params

    Returns
    -------
    Does not return anything.
    """

    def __init__(self, spaceOut, windowOut):
        self.spaces = []
        for spaceParam in spaceOut:
            # Map space parameters to Space class arguments
            longName = spaceParam[0]
            spaceId = spaceParam[1]
            sx = spaceParam[2]
            sy = spaceParam[3]
            sz = spaceParam[4]
            windows = []
            for windowParam in windowOut:
                # If window belongs to space, append the window's parameters
                if windowParam[1] == spaceParam[1]:
                    spaceParam.append(windowParam[2:])
                    for i in range(2, (len(windowParam))):
                        # The first two items are Space Name and Space Code, these are skipped
                        if i != 0 or i != 1:
                            # Map the window parameters to Window class arguments and append to the windows list
                            windowName = windowParam[i][0]
                            wx = windowParam[i][1]
                            wy = windowParam[i][2]
                            sillHeight = windowParam[i][3]
                            if sillHeight == None:
                                sillHeight = 0.1

                            wall_name = windowParam[i][4]
                            wall_length = windowParam[i][5]
                            loc_x = windowParam[i][6]
                            loc_y = windowParam[i][7]
                            windows.append(
                                Window(
                                    windowName,
                                    wx,
                                    wy,
                                    sillHeight,
                                    wall_name,
                                    wall_length,
                                    loc_x,
                                    loc_y,
                                )
                            )
                else:
                    pass

            self.spaces.append(Space(longName, spaceId, windows, sx, sy, sz))

    def print(self, spacename):
        """ Print formatted contents of Spaces for the user chosen space and its windows"""
        for space in self.spaces:
            if space.name == spacename:
                print(
                    "\n Space: \n"
                    + "\n #### Name:%s | Code:%s | Width:%0.2f | Depth:%0.2f | Height:%0.1f"
                    % (space.longName, space.name, space.sx, space.sy, space.sz)
                )
                for window in space.windows:
                    print(
                        "\n Window: \n"
                        + "Window tag: %s | Width: %0.2f | Height: %0.2f | Sill height: %0.2f \n Parent wall name: %s | Parent wall length: %0.2f | Window x location on wall: %0.2f | Window y location on wall: %0.2f"
                        % (
                            window.name,
                            window.wx,
                            window.wy,
                            window.sillHeight,
                            window.wall_name,
                            window.wall_length,
                            window.loc_x,
                            window.loc_y,
                        )
                    )


# FunSpaces = Spaces(spaceOut, windowOut)
# FunSpaces.print("A203")


def resultsOut(
    spacename,
    lightTr=0.6,
    gridsize=0.5,
    analysisPlaneHeight=0.5,
    printInfo=False,
    showPlot=False,
    blur=False,
):
    """
    Processes results into a percentage of area over 210 lux.
    Prints space info and plots results if triggered.

    Parameters
    ----------
    spacename: (str) name of the space for analysis
    lightTr: (float) visible light transmittance value
    gridsize: (float) gridsize value
    analysisPlaneHeight: (float) analysis plane height value
    printInfo: (bool) print info about space?
    showPlot: (bool) show plot?
    blur: (bool) no blur or gaussian blur on plot?

    Returns
    -------

    """
    resultList = []
    # Create an instance of Analysis object
    MyAnalysis = Analysis(
        spaceOut,
        windowOut,
        spacename,
        lightTr,
        gridsize,
        analysisPlaneHeight,
    )
    # Create an instance of space for analysis using Spaces
    AnalysisSpace = Spaces(spaceOut, windowOut)

    # Get percentace of points over 210 lux
    count = 0
    countT = 0
    # DAYLIGHT FACTOR RESULT
    for value in MyAnalysis.returnAnalysis().combined_value_by_id():
        # print("illuminance value: %d lux" % value[0])
        resultList.append(value[0])
        DFpoint = value[0] / 10000 * 100
        if DFpoint >= 2.1:
            count += 1
        countT += 1
    DFresult = count / countT * 100
    # Print the percentagewise results
    print(
        "\n"
        + 40 * "#"
        + "\n"
        + "Daylight simulation results for "
        + spacename
        + ": \n"
        + 40 * "#"
        + "\n \n"
        + str(round(DFresult, 2))
        + " % of the room has at least 210 lux"
    )
    if DFresult >= 50:
        print(
            "This room passes the Daylight Factor legislation of 2.1 %. \n \n"
            + 40 * "#"
        )
    else:
        print(
            "This room does not pass the Daylight Factor legislation of at least 2.1% DF in half of the room area.\n \n"
            + 40 * "#"
        )

    # PRINT formatted contents of space and its windows if chosen
    if printInfo is True:
        print(AnalysisSpace.print(spacename))

    # PLOT the graph of space if chosen
    if showPlot == True:
        # Find width of space analysed
        spacedim = 0
        for space in AnalysisSpace.spaces:
            if space.name == spacename:
                spacedim = float(space.sx)

        # resultList is an unnested list of results for each gridpoint and has
        # to be turned into a matrix with the correct width x depth gridpoints.
        # Get number of gridpoints on space width
        dim = int(float(spacedim) // gridsize)
        # Map resultList into a matrix using space width
        z = [resultList[i : i + dim] for i in range(0, len(resultList), dim)]

        # Set blur property
        if blur == False:
            value = "nearest"
        elif blur == True:
            value = "gaussian"

        # matplotlib stuff
        c = plt.imshow(
            z,
            cmap="YlOrRd",
            interpolation=value,
            origin="lower",
        )
        cbar = plt.colorbar(c)
        cbar.set_label("LUX")

        plt.suptitle("Daylight Analysis", fontweight="bold")

        plt.show()


######################
### USER INTERFACE ###
######################

print(
    "\n"
    + 50 * "#"
    + "\nHello! Welcome to Daylight Factor Check. \nYou can choose a space for analysis from the list below:\n"
    + 50 * "#"
)
AnalysisSpace = Spaces(spaceOut, windowOut)
spaceNameList = []
for space in AnalysisSpace.spaces:
    print("\n Space:" + "\n #### Name:%s | Code:%s" % (space.longName, space.name))
    spaceNameList.append(str(space.name))

spacename = input(
    "\n Type the code of the space you want to analyze. \n Obs! Only bedrooms, living rooms and kitchens have windows! \n Space code: \n"
)
while spacename not in spaceNameList:
    print("Ups! Wrong space code entered. Try entering a space code from the list.")
    spacename = input("Type the code of the space you want to analyze: \n")

lightTr = float(
    input(
        "Enter the visible light transmittance (VLT) for windows. Standard is 0.6.\nVLT: \n"
    )
)

gridsize = float(input("Enter the gridsize for analysis. Standard is 0.2 m: \n"))
analysisPlaneHeight = float(
    input(
        "Enter the analysis plane height for analysis. Standard is 0.5 m: \n(0.5 m is normal for residential spaces, while 0.75 m for office spaces)\nAnalysis plane height: "
    )
)
printInfo = input("Do you want to print info about space and its windows? (yes/no)\n")

while printInfo.lower() not in {"yes", "no"}:
    printInfo = input("Please enter yes or no: ")
if printInfo.lower() == "yes":
    printInfo = True
elif printInfo.lower() == "no":
    printInfo = False


showPlot = input("Do you want to show plot of analysis? (yes/no)\n")

while showPlot.lower() not in {"yes", "no"}:
    showPlot = input("Please enter yes or no: ")
if showPlot.lower() == "yes":
    showPlot = True
elif showPlot.lower() == "no":
    showPlot = False


blur = input("Do you want the grid contours to be blurred? (yes/no)\n")
while blur.lower() not in {"yes", "no"}:
    blur = input("Please enter yes or no: ")
if blur.lower() == "yes":
    blur = True
elif blur.lower() == "no":
    blur = False

print(
    resultsOut(
        spacename, lightTr, gridsize, analysisPlaneHeight, printInfo, showPlot, blur
    )
)
