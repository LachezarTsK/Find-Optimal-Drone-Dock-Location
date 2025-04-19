
using ExcelDataReader;
using NPOI.SS.Formula.Functions;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Reflection.PortableExecutable;
using System.Runtime.CompilerServices;
using System.Runtime.ConstrainedExecution;
using System.Text;
using Xunit;


namespace DroneDock
{
    public class OptimalDroneDockLocation
    {
        public enum LocationType
        {
            FIELD, BUILDING, UNIDENTIFIED
        }

        public record Point(double Latitude, double Longitude, LocationType LocationType);

        /*
         * The data that results from running the application is contained in object 'SurveyStatistics'.
         * 
         * Apart from the data in 'SurveyStatistics', the application creates output files with 
         * coordinates for the purposes of path and drone dock visualization on a map. 
         * In this respect, refer to method 'WriteFlightPathsCoordinatesToExcelFile(LocationType locationType)'
         */
        public record SurveyStatistics(string ErrorMessage,
                                       Point OptimalDroneDock,
                                       int TargetNumberOfFieldsToSurvey,
                                       int NumberOfSurveyedFields,
                                       double TotalFlightTimeInMinutes,
                                       double TotalFlightAndChargingTimeInMinutes,
                                       double DistanceInMetersToNearestBuilding,
                                       List<List<Point>> OptimalFlightPaths);

        private static readonly List<List<Point>> NO_FLIGHT_PATHS = [];

        private static readonly string NO_ERRORS = "";
        private static readonly string ERROR = "There is no input or the input is invalid!";
        private static readonly string OUTPUT_COORDINATES_START_FROM_FIELD = "DroneDock\\Output_Coordinates_Start_From_Field.xlsx";
        private static readonly string OUTPUT_COORDINATES_START_FROM_BUILDING = "DroneDock\\Output_Coordinates_Start_From_Building.xlsx";

        private static readonly double FILED_SIDE_IN_METERS = 200;
        private static readonly double CHARGING_TIME_IN_MINUTES = 35;
        private static readonly double MAXIMUM_FLIGHT_TIME_IN_MINUTES = 40;
        private static readonly double SURVEY_OF_FIELD_TIME_IN_MINUTES = 10;
        private static readonly double SPEED_BETWEEN_FIELDS_IN_METERS_PER_SECOND = 15;
        private static readonly double NEAREST_BUILDING_AND_START_ARE_ONE_AND_THE_SAME_POINT = 0;

        private readonly List<Point> fields;
        private readonly List<Point> buildings;

        private readonly List<List<Point>> flightPathsStartingFromField;
        private readonly List<List<Point>> flightPathsStartingFromBuilding;

        private readonly Point optimalCenterField;
        private readonly Point optimalCenterBuilding;

        private readonly double totalFlightTimeInMinutesStaringFromField;
        private readonly double totalFlightAndChargingTimeInMinutesStaringFromField;

        private readonly double totalFlightTimeInMinutesStaringFromBuilding;
        private readonly double totalFlightAndChargingTimeInMinutesStaringFromBuilding;

        private readonly int numberOfSurveyedFieldsStaringFromField;
        private readonly int numberOfSurveyedFieldsStaringFromBuilding;

        /*
         * Introduced for testing. Applied in class 'OptimalDroneDockLocationTests', test method
         * 'CheckCalculationOfDistancesInMetersBetweenTwoCoordinates()'
         */
        public OptimalDroneDockLocation() { }

        /*
         * Introduced for testing. Applied in class 'OptimalDroneDockLocationTests', test methods 
         * 'SelectedDroneDockForFieldHaveOptimalFlightTime()' 
         * and 
         * 'SelectedDroneDockForBuildingHaveOptimalFlightTime()'
         */
        public OptimalDroneDockLocation(List<Point> fields,
                                        List<Point> buildings,
                                        Point centerField,
                                        Point centerBuilding)
        {
            this.fields = fields;
            this.buildings = buildings;

            optimalCenterField = centerField;
            optimalCenterBuilding = centerBuilding;

            flightPathsStartingFromField = CreateOptimalFlightPathsFromSelectedCenter(LocationType.FIELD);
            flightPathsStartingFromBuilding = CreateOptimalFlightPathsFromSelectedCenter(LocationType.BUILDING);

            totalFlightTimeInMinutesStaringFromField = GetTotalFlightTimeInMinutes(LocationType.FIELD);
            totalFlightAndChargingTimeInMinutesStaringFromField = GetTotalFlightAndChargingTimeInMinutes(LocationType.FIELD);

            totalFlightTimeInMinutesStaringFromBuilding = GetTotalFlightTimeInMinutes(LocationType.BUILDING);
            totalFlightAndChargingTimeInMinutesStaringFromBuilding = GetTotalFlightAndChargingTimeInMinutes(LocationType.BUILDING);

            numberOfSurveyedFieldsStaringFromField = GetNumberOfSurveyedFields(LocationType.FIELD);
            numberOfSurveyedFieldsStaringFromBuilding = GetNumberOfSurveyedFields(LocationType.BUILDING);
        }

        /*
         * The actual constructor that runs the application.
         */
        public OptimalDroneDockLocation(string filePath)
        {
            fields = ExtractCoordinatesFromFile(filePath, LocationType.FIELD);
            buildings = ExtractCoordinatesFromFile(filePath, LocationType.BUILDING);

            optimalCenterField = FindCenterLocationForMinTotalDistanceToAllFields(LocationType.FIELD);
            optimalCenterBuilding = FindCenterLocationForMinTotalDistanceToAllFields(LocationType.BUILDING);

            flightPathsStartingFromField = CreateOptimalFlightPathsFromSelectedCenter(LocationType.FIELD);
            flightPathsStartingFromBuilding = CreateOptimalFlightPathsFromSelectedCenter(LocationType.BUILDING);

            totalFlightTimeInMinutesStaringFromField = GetTotalFlightTimeInMinutes(LocationType.FIELD);
            totalFlightAndChargingTimeInMinutesStaringFromField = GetTotalFlightAndChargingTimeInMinutes(LocationType.FIELD);

            totalFlightTimeInMinutesStaringFromBuilding = GetTotalFlightTimeInMinutes(LocationType.BUILDING);
            totalFlightAndChargingTimeInMinutesStaringFromBuilding = GetTotalFlightAndChargingTimeInMinutes(LocationType.BUILDING);

            numberOfSurveyedFieldsStaringFromField = GetNumberOfSurveyedFields(LocationType.FIELD);
            numberOfSurveyedFieldsStaringFromBuilding = GetNumberOfSurveyedFields(LocationType.BUILDING);

            WriteFlightPathsCoordinatesToExcelFile(LocationType.FIELD);
            WriteFlightPathsCoordinatesToExcelFile(LocationType.BUILDING);
        }

        /*
         * Create a list of two objects 'SurveyStatistics'.
         * The first one contains tha data for a drone dock location on a field.
         * The second one contains tha data for a drone dock location on a building.
         */
        public List<SurveyStatistics> GetSurveyStatistics()
        {
            SurveyStatistics optionStartFromField;
            SurveyStatistics optionStartFromBuilding;

            if (fields == null || fields.Count == 0)
            {
                optionStartFromField = new SurveyStatistics(ERROR, null, 0, 0, 0, 0, 0, null);
            }
            else
            {
                optionStartFromField = new SurveyStatistics
                    (NO_ERRORS,
                     optimalCenterField,
                     fields.Count,
                     numberOfSurveyedFieldsStaringFromField,
                     totalFlightTimeInMinutesStaringFromField,
                     totalFlightAndChargingTimeInMinutesStaringFromField,
                     GetDistanceInMetersFromCenterFieldToNearestBuilding(),
                     flightPathsStartingFromField);
            }

            if (fields == null || fields.Count == 0 || buildings == null || buildings.Count == 0)
            {
                optionStartFromBuilding = new SurveyStatistics(ERROR, null, 0, 0, 0, 0, 0, null);
            }
            else
            {
                optionStartFromBuilding = new SurveyStatistics
                    (NO_ERRORS,
                     optimalCenterBuilding,
                     fields.Count,
                     numberOfSurveyedFieldsStaringFromBuilding,
                     totalFlightTimeInMinutesStaringFromBuilding,
                     totalFlightAndChargingTimeInMinutesStaringFromBuilding,
                     NEAREST_BUILDING_AND_START_ARE_ONE_AND_THE_SAME_POINT,
                     flightPathsStartingFromBuilding);
            }


            return [optionStartFromField, optionStartFromBuilding];
        }

        /*
         * The input file containing the coordinates is expected to be in Excel.
         * Both latitude and longitude are expected to be in the first column.
         * The type of location, i.e. either field or building, is expected to be in the second column. 
         * Example:
         *
         *            First Column                Second Column
         * 
         * 49.15868902252248, 9.111073073485683	      Field
         * 49.1582580129513, 9.113612777241652	      Building
         * etc.
         */
        private static List<Point> ExtractCoordinatesFromFile(string filePath, LocationType locationType)
        {
            int columnForCoordinates = 0;
            int columnForDescription = 1;
            List<Point> points = [];

            FileStream stream;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            try
            {
                stream = File.OpenRead(filePath);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                return points;
            }


            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    while (reader.Read() && reader.FieldCount >= 1 + columnForCoordinates && reader.GetValue(columnForCoordinates) != null)
                    {
                        string[] coordinates = reader.GetValue(columnForCoordinates).ToString().Split(",");

                        try
                        {
                            double latitude = double.Parse(coordinates[0]);
                            double longitude = double.Parse(coordinates[1]);

                            if (reader.FieldCount >= 1 + columnForDescription
                                && reader.GetValue(columnForDescription) != null
                                && locationType == GetLocationType(reader.GetValue(columnForDescription).ToString()))
                            {
                                points.Add(new Point(latitude, longitude, locationType));
                            }
                        }
                        catch (Exception e)
                        {
                            Console.Error.WriteLine(e.Message);
                        }

                    }
                } while (reader.NextResult());
            }
            stream.Close();

            return points;
        }

        /*
         *The output file is in Excel. 
         *The first column contains latitude coordinates. 
         *The second column contains longitude coordinates.
         *          
         * Since the visualization is done at https://www.gpsvisualizer.com/map?output_google 
         * to visualize all flights per survey with different colors (when visualizing all flights per survey at once) 
         * the row preceding the coordinates for each flight has ‘latitude’ as title in the first column
         * and ‘longitude’ as title in second column. Example:
         *
         * First Column   Second Column
         *   
         *  latitude	    longitude
         * 49.15742599	   9.116101442
         * 49.15621042	   9.116235216
         * … the rest of the coordinates for first flight of current survey …
         * latitude	       longitude
         * … coordinates for second flight of current survey …
         * etc.
         */
        private void WriteFlightPathsCoordinatesToExcelFile(LocationType locationType)
        {
            List<List<Point>> flightPaths = (locationType == LocationType.FIELD)
                                            ? flightPathsStartingFromField
                                            : flightPathsStartingFromBuilding;

            if (flightPaths.Count == 0)
            {
                return;
            }

            string filePath = (locationType == LocationType.FIELD)
                                            ? OUTPUT_COORDINATES_START_FROM_FIELD
                                            : OUTPUT_COORDINATES_START_FROM_BUILDING;

            var book = new XSSFWorkbook();
            var sheet = book.CreateSheet("sheet1");
            int firstColumn = 0;
            int secondColumn = 1;

            int line = 0;
            foreach (var path in flightPaths)
            {
                var row = sheet.CreateRow(line);
                row.CreateCell(firstColumn).SetCellValue("latitude");
                row.CreateCell(secondColumn).SetCellValue("longitude");
                ++line;

                foreach (Point point in path)
                {
                    row = sheet.CreateRow(line);
                    row.CreateCell(firstColumn).SetCellValue(point.Latitude);
                    row.CreateCell(secondColumn).SetCellValue(point.Longitude);
                    ++line;
                }
            }

            try
            {
                using var fileStream = File.Create(filePath);
                book.Write(fileStream);
                fileStream.Close();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
            }
            book.Close();
        }

        /*
         * Finds the optimum drone dock for the given type of location.
         * In this application, the valid type of location for a drone dock is either 'field' or 'building'.
         * 
         * Optimal Drone Dock: the point from where minimum flight time can be achieved
         *                     to survey the fields, under current constraints.
         * 
         * It is assumed that the start of the survey for each field is from the center and that it is 
         * completed also at the center. Each field must be surveyed at once, during one flight.
         */
        private Point FindCenterLocationForMinTotalDistanceToAllFields(LocationType locationType)
        {

            if (fields.Count == 0 || (locationType == LocationType.BUILDING && buildings.Count == 0))
            {
                return new Point(0, 0, LocationType.UNIDENTIFIED);
            }

            Point optimalCenter = new(0, 0, locationType);
            List<Point> points = (locationType == LocationType.FIELD) ? fields : buildings;

            double minFlightTimeForSurvey = double.MaxValue;
            int maxNumberSurveyedFieldsFromCenter = 0;

            foreach (var center in points)
            {
                int numberOfSurveyedFieldsFromCurrentCenter = 0;
                double flightTimeForSurveyedFieldsFromCurrentCenter = 0;
                HashSet<Point> fieldsToSurvey = [.. fields];

                while (true)
                {
                    /*
                     * checking 'fieldsToSurvey.Count > 0' is not enough, since there might be fields at a distance,
                     * which requires more time than the maximum flight time of the drone, therefore, in such cases
                     * 'fieldsToSurvey.Count' will always be greater than '0', and the loop will become infinite.
                     */
                    bool newFieldsSurveyed = false;
                    double currentFlightTimeInMinutes = 0;
                    Point current = center;

                    while (true)
                    {
                        Point next = GetUnsurveyedFieldAtMinDistance(current, fieldsToSurvey);
                        double timeInSecondsFromCurrentToNext = GetTimeInSecondsToFlightBetweenTwoPoints(current, next);
                        double timeInSecondsFromNextToCenter = GetTimeInSecondsToFlightBetweenTwoPoints(next, center);

                        if (HasTimeToSurveyNextFieldAndReturnToCenter(timeInSecondsFromCurrentToNext,
                                                                      timeInSecondsFromNextToCenter,
                                                                      currentFlightTimeInMinutes))
                        {
                            currentFlightTimeInMinutes += ConvertToMinutes(timeInSecondsFromCurrentToNext) + SURVEY_OF_FIELD_TIME_IN_MINUTES;
                            fieldsToSurvey.Remove(next);
                            newFieldsSurveyed = true;
                            ++numberOfSurveyedFieldsFromCurrentCenter;
                        }
                        else
                        {
                            double timeInSecondsFromCurrentToCenter = GetTimeInSecondsToFlightBetweenTwoPoints(current, center);
                            currentFlightTimeInMinutes += ConvertToMinutes(timeInSecondsFromCurrentToCenter);
                            flightTimeForSurveyedFieldsFromCurrentCenter += currentFlightTimeInMinutes;
                            break;
                        }
                        current = next;
                    }

                    if (!newFieldsSurveyed)
                    {
                        break;
                    }
                }

                /*
                 * If not all fields are reachable from current center, i.e. these are beyond the maximum flight time,
                 * select as center the point from where most fields are reachable, therefore: 
                 * 'maxNumberSurveyedFieldsFromCenter < numberOfSurveyedFieldsFromCurrentCenter'
                 */
                if (maxNumberSurveyedFieldsFromCenter < numberOfSurveyedFieldsFromCurrentCenter || minFlightTimeForSurvey > flightTimeForSurveyedFieldsFromCurrentCenter)
                {
                    optimalCenter = center;
                    minFlightTimeForSurvey = flightTimeForSurveyedFieldsFromCurrentCenter;
                    maxNumberSurveyedFieldsFromCenter = numberOfSurveyedFieldsFromCurrentCenter;
                }
            }
            return optimalCenter;
        }

        /*
         * Creates the optimal flight paths on the basis of the selected optimal drone dock for the given
         * type of location.
         * 
         * Optimal Flight Path: the path with minimum flight to survey the fields, under current constraints.
         * 
         * It is assumed that the start of the survey for each field is from the center and that it is 
         * completed also at the center. Each field must be surveyed at once, during one flight.
         */
        private List<List<Point>> CreateOptimalFlightPathsFromSelectedCenter(LocationType locationType)
        {
            Point center = (locationType == LocationType.FIELD)
                 ? optimalCenterField
                 : optimalCenterBuilding;

            if (center.LocationType == LocationType.UNIDENTIFIED || fields.Count == 0)
            {
                return NO_FLIGHT_PATHS;
            }

            List<List<Point>> flightPaths = [];
            HashSet<Point> fieldsToSurvey = [.. fields];

            while (true)
            {
                /*
                 * If not all fields are reachable from selected optimal center, i.e. these are beyond the maximum flight time,
                 * 'fieldsToSurvey.Count > 0' will always be true, therefore, implementing 'bool newFieldsSurveyed'
                 */
                bool newFieldsSurveyed = false;
                double currentFlightTimeInMinutes = 0;
                Point current = center;

                List<Point> flightPathPerOneCharge = [];
                flightPathPerOneCharge.Add(current);

                while (true)
                {
                    Point next = GetUnsurveyedFieldAtMinDistance(current, fieldsToSurvey);
                    double timeInSecondsFromCurrentToNext = GetTimeInSecondsToFlightBetweenTwoPoints(current, next);
                    double timeInSecondsFromNextToCenter = GetTimeInSecondsToFlightBetweenTwoPoints(next, center);

                    if (HasTimeToSurveyNextFieldAndReturnToCenter(timeInSecondsFromCurrentToNext,
                                                                  timeInSecondsFromNextToCenter,
                                                                  currentFlightTimeInMinutes))
                    {
                        currentFlightTimeInMinutes += ConvertToMinutes(timeInSecondsFromCurrentToNext) + SURVEY_OF_FIELD_TIME_IN_MINUTES;
                        fieldsToSurvey.Remove(next);
                        flightPathPerOneCharge.Add(next);
                        newFieldsSurveyed = true;
                    }
                    else
                    {
                        double timeInSecondsFromCurrentToCenter = GetTimeInSecondsToFlightBetweenTwoPoints(current, center);
                        currentFlightTimeInMinutes += ConvertToMinutes(timeInSecondsFromCurrentToCenter);
                        flightPathPerOneCharge.Add(center);
                        break;
                    }
                    current = next;
                }

                if (!newFieldsSurveyed)
                {
                    break;
                }
                flightPaths.Add(flightPathPerOneCharge);
            }
            return flightPaths;
        }

        /*
         * Helper method, applied in 'CreateOptimalFlightPathsFromSelectedCenter(LocationType locationType).
         * It finds the next field to survey.
         */
        private static Point GetUnsurveyedFieldAtMinDistance(Point from, HashSet<Point> fieldsToSurvey)
        {
            Point next = new(0, 0, LocationType.FIELD);
            double minDistance = double.MaxValue;

            foreach (Point to in fieldsToSurvey)
            {
                double distance = GetDistanceInMeters(from, to);
                if (minDistance > distance)
                {
                    minDistance = distance;
                    next = to;
                }
            }
            return next;
        }

        /*
        * Helper method, applied in 'FindCenterLocationForMinTotalDistanceToAllFields(LocationType locationType)'
        * and in 'CreateOptimalFlightPathsFromSelectedCenter(LocationType locationType)'.
        * Checks whether there is time to survey the next field and then to go back to the drone dock.
        */
        private static bool HasTimeToSurveyNextFieldAndReturnToCenter(double timeInSecondsFromCurrentToNext,
                                                               double timeInSecondsFromNextToCenter,
                                                               double currentFlightTimeInMinutes)
        {
            return ConvertToMinutes(timeInSecondsFromCurrentToNext + timeInSecondsFromNextToCenter)
                        + currentFlightTimeInMinutes + SURVEY_OF_FIELD_TIME_IN_MINUTES
                        <= MAXIMUM_FLIGHT_TIME_IN_MINUTES;
        }

        private static double GetTimeInSecondsToFlightBetweenTwoPoints(Point first, Point second)
        {
            return GetDistanceInMeters(first, second) / SPEED_BETWEEN_FIELDS_IN_METERS_PER_SECOND;
        }

        /*
         * The method is applied in 'GetSurveyStatistics()' and is just for the sake of giving 
         * as much useful information as possible. And when selecting a drone dock,
         * it would be useful to know not just the optimal drone dock for a field and for a building 
         * but also the distance from the optimal drone dock for a field to the nearest building 
         * (which, of course, might or might not coincide with the optimal drone dock for a building).
         */
        private double GetDistanceInMetersFromCenterFieldToNearestBuilding()
        {
            double minDistance = double.MaxValue;
            foreach (var building in buildings)
            {
                double currentDistance = GetDistanceInMeters(optimalCenterField, building);
                minDistance = Math.Min(minDistance, currentDistance);
            }
            return minDistance;
        }


        private double GetTotalFlightTimeInMinutes(LocationType locationType)
        {
            if (locationType == LocationType.UNIDENTIFIED)
            {
                return 0;
            }

            double flightTimeInSeconds = 0;

            List<List<Point>> paths = (locationType == LocationType.FIELD)
                                       ? flightPathsStartingFromField
                                       : flightPathsStartingFromBuilding;

            foreach (var path in paths)
            {
                for (int i = 0; i < path.Count - 1; ++i)
                {
                    flightTimeInSeconds += GetTimeInSecondsToFlightBetweenTwoPoints(path[i], path[i + 1]);
                }
            }

            return ConvertToMinutes(flightTimeInSeconds) + GetNumberOfSurveyedFields(locationType) * SURVEY_OF_FIELD_TIME_IN_MINUTES;
        }

        /*
         * Time for charging before the first flight and after the last flight is not included.
         */
        private double GetTotalFlightAndChargingTimeInMinutes(LocationType locationType)
        {
            if (locationType == LocationType.UNIDENTIFIED)
            {
                return 0;
            }

            double totalFlightTimeInMinutes = (locationType == LocationType.FIELD)
                                       ? totalFlightTimeInMinutesStaringFromField
                                       : totalFlightTimeInMinutesStaringFromBuilding;

            int numberOfFlights = (locationType == LocationType.FIELD)
                                       ? flightPathsStartingFromField.Count
                                       : flightPathsStartingFromBuilding.Count;

            return totalFlightTimeInMinutes + (numberOfFlights - 1) * CHARGING_TIME_IN_MINUTES;
        }

        /*
         * Helper method, used in 'GetTotalFlightTimeInMinutes(LocationType locationType)'
         * and in 'GetSurveyStatistics()'
         */
        private int GetNumberOfSurveyedFields(LocationType locationType)
        {
            List<List<Point>> paths = (locationType == LocationType.FIELD)
                ? flightPathsStartingFromField : flightPathsStartingFromField;

            int countSurveyedFields = 0;
            foreach (var path in paths)
            {
                countSurveyedFields += path.Count;
            }
            int dockingPointAtStartAndEndOfEachRoundOfFlight = paths.Count * 2;
            return countSurveyedFields - dockingPointAtStartAndEndOfEachRoundOfFlight;
        }

        public List<Point> GetFields()
        {
            return [.. fields];
        }

        public List<Point> GetBuildings()
        {
            return [.. buildings];
        }

        /*
         * To have a better precision during the calculations, time in seconds is used.
         * And for the presentation of the results, the seconds are converted to minutes.
         * Therefore, the current method.
         */
        private static double ConvertToMinutes(double seconds)
        {
            return seconds / 60;
        }

        private static double DegreesToRadians(double degrees)
        {
            return degrees * Math.PI / 180;
        }

        /*
         * Calculates the distance between the coordinates of two GPS points.
         */
        private static double GetDistanceInMeters(Point first, Point second)
        {
            double radiusEarthInKilometers = 6371;
            double dLat = DegreesToRadians(first.Latitude - second.Latitude);
            double dLon = DegreesToRadians(first.Longitude - second.Longitude);

            double lat1 = DegreesToRadians(first.Latitude);
            double lat2 = DegreesToRadians(second.Latitude);

            double a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) + Math.Sin(dLon / 2) * Math.Sin(dLon / 2) * Math.Cos(lat1) * Math.Cos(lat2);
            double c = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));

            return radiusEarthInKilometers * c * 1000;
        }

        /*
         * Helper method, applied in 'ExtractCoordinatesFromFile(string filePath, LocationType locationType)'.
         * Utilized when reading the input file, in order to determine to which type of location the coordinates belong.
         */
        private static LocationType GetLocationType(string description)
        {
            if (description.Trim().ToLower().Equals("field"))
            {
                return LocationType.FIELD;
            }
            if (description.Trim().ToLower().Equals("building"))
            {
                return LocationType.BUILDING;
            }
            return LocationType.UNIDENTIFIED;
        }

        /*
         * Presentation of the results.
         */
        public void PrintSurveyStatistics()
        {
            int surveyNo = 1;

            foreach (var current in GetSurveyStatistics())
            {
                Console.WriteLine("----- START STATISTICS FOR SURVEY NO " + surveyNo + " -----");

                if (current.ErrorMessage.Length > 0)
                {
                    Console.WriteLine(current.ErrorMessage);
                    Console.WriteLine("----- END STATISTICS FOR SURVEY NO " + surveyNo + " -----");
                    Console.WriteLine("\n");
                    ++surveyNo;
                    continue;
                }


                Console.WriteLine();
                Console.WriteLine("Optimal Drone Dock: " + current.OptimalDroneDock);
                Console.WriteLine("Target Number Of Fields To Survey: " + current.TargetNumberOfFieldsToSurvey);
                Console.WriteLine("Number Of Surveyed Fields: " + current.NumberOfSurveyedFields);
                Console.WriteLine("Total Flight Time In Minutes: " + Double.Round(current.TotalFlightTimeInMinutes, 2));
                Console.WriteLine("Total Flight And Charging Time In Minutes: " + Double.Round(current.TotalFlightAndChargingTimeInMinutes, 2));
                Console.WriteLine("Distance In Meters To Nearest Building: " + Double.Round(current.DistanceInMetersToNearestBuilding, 2));
                Console.WriteLine();

                int flightNo = 1;
                foreach (var path in current.OptimalFlightPaths)
                {
                    Console.WriteLine("Survey No: " + surveyNo + ", Flight No: " + flightNo);
                    foreach (var point in path)
                    {
                        Console.WriteLine(point);
                    }
                    Console.WriteLine();
                    ++flightNo;
                }

                Console.WriteLine("----- END STATISTICS FOR SURVEY NO " + surveyNo + " -----");
                Console.WriteLine("\n");
                ++surveyNo;
            }
        }
    }
}



public class Program
{
    static void Main()
    {
        string filePath = "DroneDock\\Input_Coordinates_Fields_and_Buildings.xlsx";
        var droneDock = new DroneDock.OptimalDroneDockLocation(filePath);
        droneDock.PrintSurveyStatistics();
    }
}
