
using DroneDock;
using MathNet.Numerics.Integration;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Util;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Xunit.Sdk;
using static DroneDock.OptimalDroneDockLocation;
using static NPOI.HSSF.UserModel.HeaderFooter;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace DroneDock.UnitTests
{
    public class OptimalDroneDockLocationTests
    {
        private record Distance(Point First, Point Second, double MetersBetweenPoints);

        /*
         * Source of 'expectedDistances' https://www.movable-type.co.uk/scripts/latlong.html
         */
        private static readonly List<Distance> expectedDistances =
        [
            new Distance(new Point(25.246013772085803, 25.246013772085803, LocationType.UNIDENTIFIED),
                         new Point(25.246013772085803, 25.246013772085803, LocationType.UNIDENTIFIED),
                         ConvertToMeters(0.000)),

            new Distance(new Point(48.36810950039437, 8.303569430210762, LocationType.UNIDENTIFIED),
                         new Point(48.370145411609506, 8.527438383780646, LocationType.UNIDENTIFIED),
                         ConvertToMeters(16.54)),

            new Distance(new Point(51.75808602920108, 12.928039915061246, LocationType.UNIDENTIFIED),
                         new Point(51.78483760192261, 12.94131467183842, LocationType.UNIDENTIFIED),
                         ConvertToMeters(3.112)),

            new Distance(new Point(43.86260284936179, -84.63478251487683, LocationType.UNIDENTIFIED),
                         new Point(43.86275502600581, -84.63447860450665, LocationType.UNIDENTIFIED),
                         ConvertToMeters(0.02966)),

             new Distance(new Point(-25.246013772085803, 130.99454600534904, LocationType.UNIDENTIFIED),
                          new Point(-25.246006449757466, 130.99464314650047, LocationType.UNIDENTIFIED),
                          ConvertToMeters(0.009804)),

              new Distance(new Point(-30.374654700590366, 131.49545065392697, LocationType.UNIDENTIFIED),
                           new Point(-30.148451741719285, 131.94143977560225, LocationType.UNIDENTIFIED),
                           ConvertToMeters(49.67)),

             new Distance(new Point(35.08907695222524, 135.58034444637707, LocationType.UNIDENTIFIED),
                          new Point(34.932293832738104, 135.81041145288606, LocationType.UNIDENTIFIED),
                          ConvertToMeters(27.26)),

              new Distance(new Point(0.2207436038771777, 21.806332748154157, LocationType.UNIDENTIFIED),
                           new Point(0.4885481453619999, 22.10227223631869, LocationType.UNIDENTIFIED),
                           ConvertToMeters(44.38))
        ];

        private static readonly double EPSILON_IN_METERS = ConvertToMeters(0.02);

        private static readonly string NO_ERRORS = "";
        private static readonly string ERROR = "There is no input or the input is invalid!";

        private static readonly SurveyStatistics errorSurveyStatistics = new(ERROR, null, 0, 0, 0, 0, 0, null);

        private static readonly string FILE_PATH_COMPLETE_INPUT_COORDINATES_FIELDS_AND_BUILDINGS =
                                                             "\\DroneDock.UnitTests\\CompleteInput_Coordinates_Fields_and_Buildings.xlsx";
        private static readonly string FILE_PATH_INPUT_COORDINATES_WHERE_NOT_ALL_FIELDS_ARE_REACHABLE =
                                                             "\\DroneDock.UnitTests\\Input_Coordinates_Where_Not_All_Fields_Are_Reachable.xlsx";
        private static readonly string FILE_PATH_INPUT_COORDINATES_ONLY_FOR_FIELDS =
                                                              "\\DroneDock.UnitTests\\Input_Coordinates_Only_For_Fields.xlsx";
        private static readonly string FILE_PATH_INPUT_COORDINATES_ONLY_FOR_BUILDINGS =
                                                               "\\DroneDock.UnitTests\\Input_Coordinates_Only_For_Buildings.xlsx";
        private static readonly string FILE_PATH_EMPTY_INPUT = "\\DroneDock.UnitTests\\Empty_Input.xlsx";
        private static readonly string FILE_PATH_NONEXISTENT = "\\NONEXISTENT_FILE_PATH";


        private static double ConvertToMeters(double kilometers)
        {
            return kilometers * 1000;
        }

        /*
         * Checks that if the input file is null, the there is no valid output at all.
         */
        [Fact]
        public void InvalidInput_NonexistentFilePath()
        {
            string filePath = FILE_PATH_NONEXISTENT;
            var droneDock = new OptimalDroneDockLocation(filePath);

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.Empty(fields);
            Assert.Empty(buildings);

            foreach (var currentSurveyStatistics in droneDock.GetSurveyStatistics())
            {
                Assert.True(currentSurveyStatistics == errorSurveyStatistics);
            }
        }

        /*
         * Checks that if the input file is empty, the there is no valid output at all.
         */
        [Fact]
        public void InvalidInput_FileIsEmpty()
        {
            string filePath = FILE_PATH_EMPTY_INPUT;
            var droneDock = new OptimalDroneDockLocation(filePath);

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.Empty(fields);
            Assert.Empty(buildings);

            foreach (var currentSurveyStatistics in droneDock.GetSurveyStatistics())
            {
                Assert.True(currentSurveyStatistics == errorSurveyStatistics);
            }
        }

        /*
         * Checks that if there is an input only for buildings, the there is no valid output at all, 
         * since without fields there is nothing to survey.
         */
        [Fact]
        public void InvalidInput_InputCoordinatesOnlyForBuildings()
        {
            string filePath = FILE_PATH_INPUT_COORDINATES_ONLY_FOR_BUILDINGS;
            var droneDock = new OptimalDroneDockLocation(filePath);

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.Empty(fields);
            Assert.True(buildings.Count > 0);

            foreach (var currentSurveyStatistics in droneDock.GetSurveyStatistics())
            {
                Assert.True(currentSurveyStatistics == errorSurveyStatistics);
            }
        }

        /*
         * Check that if there is an input only for fields, then there is a valid output only for a field drone dock.         * 
         */
        [Fact]
        public void PartialInput_InputCoordinatesOnlyForFields()
        {
            string filePath = FILE_PATH_INPUT_COORDINATES_ONLY_FOR_FIELDS;
            var droneDock = new OptimalDroneDockLocation(filePath);

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.True(fields.Count > 0);
            Assert.Empty(buildings);

            foreach (var currentSurveyStatistics in droneDock.GetSurveyStatistics())
            {
                if (currentSurveyStatistics.OptimalDroneDock != null &&
                    currentSurveyStatistics.OptimalDroneDock.LocationType == LocationType.FIELD)
                {
                    Assert.False(currentSurveyStatistics == errorSurveyStatistics);
                }
                else
                {
                    Assert.True(currentSurveyStatistics == errorSurveyStatistics);
                }
            }
        }

        /*
         * Checks that if not all fields are reachable ,i.e. within the drone max flight time 
         * (allowing for time to survey the field and go back to the drone dock), then the drone dock
         * is selected is such a way that the maximum number of fields are surveyed.
         * 
         * Example: 
         * 35 fields to be surveyed 
         * 4 fields are close to each other and can be surveyed but the other 31 can not be reached from these 4 fields.
         * 31 fields are close to each other and can be surveyed but the other 4 can not be reached from these 31 fields.
         * Then 31 fields are surveyed.
         */
        [Fact]
        public void PartialInput_InputCoordinatesWhereNotAllFieldsAreReachable()
        {
            string filePath = FILE_PATH_INPUT_COORDINATES_WHERE_NOT_ALL_FIELDS_ARE_REACHABLE;
            var droneDock = new OptimalDroneDockLocation(filePath);
            int expectedNumberOfFieldsNotSurveyed = 4;

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.True(fields.Count > 0);
            Assert.True(buildings.Count > 0);

            foreach (var currentSurveyStatistics in droneDock.GetSurveyStatistics())
            {
                int numberOfFieldsNotSurveyed = currentSurveyStatistics.TargetNumberOfFieldsToSurvey
                                                - currentSurveyStatistics.NumberOfSurveyedFields;
                Assert.Equal(expectedNumberOfFieldsNotSurveyed, numberOfFieldsNotSurveyed);
            }
        }

        /*
         * Checks that if all fields are reachable ,i.e. within the drone max flight time 
         * (allowing for time to survey the field and go back to the drone dock), then all fields are surveyed.
         */
        [Fact]
        public void CompleteInput_InputCoordinatesWhereAllFieldsAreReachable()
        {
            string filePath = FILE_PATH_COMPLETE_INPUT_COORDINATES_FIELDS_AND_BUILDINGS;
            var droneDock = new OptimalDroneDockLocation(filePath);
            int expectedNumberOfFieldsNotSurveyed = 0;

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.True(fields.Count > 0);
            Assert.True(buildings.Count > 0);

            foreach (var currentSurveyStatistics in droneDock.GetSurveyStatistics())
            {
                int numberOfFieldsNotSurveyed = currentSurveyStatistics.TargetNumberOfFieldsToSurvey
                                                - currentSurveyStatistics.NumberOfSurveyedFields;
                Assert.Equal(expectedNumberOfFieldsNotSurveyed, numberOfFieldsNotSurveyed);
            }
        }

        /*
         * Checks that the method for calculating distance between to GPS coordinates 
         * calculate correctly within the given deviation, i.e. EPSILON_IN_METERS. 
         *
         * Max flight distance (and this is without return trip and without surveying fields) 
         * is 15 meters per second x 60 seconds per minutes x 40 minutes max flight time = 36 km. 
         * Therefore, distances measured here are in the range of [0, 50] km and EPSILON for this range 
         * is 0.02 km when comparing my application and https://www.movable-type.co.uk/scripts/latlong.html. 
         * There are differences between almost all online calculators for measuring distances between 
         * GPS coordinates, as well as between these and my method, and these differences increase 
         * with the increase in distance.
         *
         * Example:
         * The distance between 
         * Point (0.2407436038771777, 21.936332748154157) and Point (0.3385481453619999, 22.40227223631869) 
         * on https://www.movable-type.co.uk/scripts/latlong.html  is 52.94 km 
         * and on https://www.calculator.net (section for GPS coordinates) it is 52.98 km, 
         * i.e. EPSILON = 0.04 km. And for coordinates Point (-90, -180) and Point (90, 180) 
         * the first site returns a distance of 20 020 km and the latter site returns a distance 
         * of 20 003.8 km, i.e. EPSILON = 16.2 km. And such is approximately the EPSILON ratio when 
         * comparing other online calculators for distances based on GPS coordinates, 
         * as well as between these and the method in my application.
         *
         * Therefore, if larger distances are to be tested, the value of EPSILON must be adjusted accordingly.
         */
        [Fact]
        public void CheckCalculationOfDistancesInMetersBetweenTwoCoordinates()
        {
            var droneDock = new OptimalDroneDockLocation();
            MethodInfo method = droneDock.GetType().GetMethod("GetDistanceInMeters", BindingFlags.NonPublic | BindingFlags.Static);

            foreach (var current in expectedDistances)
            {
                object[] parameters = [current.First, current.Second];
                double distance = (double)method!.Invoke(droneDock, parameters)!;
                Assert.True(Math.Abs((distance) - current.MetersBetweenPoints) <= EPSILON_IN_METERS);
            }
        }

        /*
         * Checks that the selected optimal drone dock on a field has minimum flight time
         * compared to all other fields that could serve as a drone dock.
         */
        [Fact]
        public void SelectedDroneDockForFieldHaveOptimalFlightTime()
        {
            string filePath = FILE_PATH_COMPLETE_INPUT_COORDINATES_FIELDS_AND_BUILDINGS;
            var droneDock = new OptimalDroneDockLocation(filePath);

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.True(fields.Count > 0);
            Assert.True(buildings.Count > 0);

            var surveyResultsForField = droneDock.GetSurveyStatistics()[0];
            var surveyResultsForBuilding = droneDock.GetSurveyStatistics()[1];

            foreach (var fieldToBeCheckedAsOptimalDroneDock in fields)
            {
                var surveyResultsToTest = new OptimalDroneDockLocation(fields,
                                                                 buildings,
                                                                 fieldToBeCheckedAsOptimalDroneDock,
                                                                 surveyResultsForBuilding.OptimalDroneDock);

                foreach (var survey in surveyResultsToTest.GetSurveyStatistics())
                {
                    if (survey.OptimalDroneDock.LocationType == LocationType.FIELD
                        && survey.NumberOfSurveyedFields == surveyResultsForField.NumberOfSurveyedFields
                        && survey.OptimalDroneDock != surveyResultsForField.OptimalDroneDock)
                    {
                        Assert.True(survey.TotalFlightTimeInMinutes > surveyResultsForField.TotalFlightTimeInMinutes);
                    }
                }
            }
        }

        /*
        * Checks that the selected optimal drone dock on a building has minimum flight time
        * compared to all other buildings that could serve as a drone dock.
        */
        [Fact]
        public void SelectedDroneDockForBuildingHaveOptimalFlightTime()
        {
            string filePath = FILE_PATH_COMPLETE_INPUT_COORDINATES_FIELDS_AND_BUILDINGS;
            var droneDock = new OptimalDroneDockLocation(filePath);

            var fields = droneDock.GetFields();
            var buildings = droneDock.GetBuildings();

            Assert.True(fields.Count > 0);
            Assert.True(buildings.Count > 0);

            var surveyResultsForField = droneDock.GetSurveyStatistics()[0];
            var surveyResultsForBuilding = droneDock.GetSurveyStatistics()[1];

            foreach (var buildingToBeCheckedAsOptimalDroneDock in buildings)
            {
                var surveyResultsToTest = new OptimalDroneDockLocation(fields,
                                                                       buildings,
                                                                       surveyResultsForField.OptimalDroneDock,
                                                                       buildingToBeCheckedAsOptimalDroneDock);

                foreach (var survey in surveyResultsToTest.GetSurveyStatistics())
                {
                    if (survey.OptimalDroneDock.LocationType == LocationType.BUILDING
                        && survey.NumberOfSurveyedFields == surveyResultsForBuilding.NumberOfSurveyedFields
                        && survey.OptimalDroneDock != surveyResultsForBuilding.OptimalDroneDock)
                    {
                        Assert.True(survey.TotalFlightTimeInMinutes > surveyResultsForBuilding.TotalFlightTimeInMinutes);
                    }
                }
            }
        }
    }
}
