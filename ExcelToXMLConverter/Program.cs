using OfficeOpenXml;
using System.Xml.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelToXMLConverter
{
    internal class Program
    {
        private static void Main()
        {
            try
            {
                // Set EPPlus license context to NonCommercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Load the XML template file into memory
                XDocument xmlTemplate;
                using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                {
                    xmlTemplate = XDocument.Load(stream);
                }

                // Load the XML seals list file into memory
                XDocument sealsList;
                using (var stream = new StreamReader(@"../all_seals.xml"))
                {
                    sealsList = XDocument.Load(stream);
                }

                XNamespace ns = "http://www.tei-c.org/ns/1.0";

                // Load the Excel file into memory
                ExcelPackage package;
                using (var stream = new FileStream(@"./resources/SIGIDOC CELLS ENG.xlsx", FileMode.Open, FileAccess.Read))
                {
                    package = new ExcelPackage(stream);
                }
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet == null)
                {
                    Console.WriteLine("The worksheet does not exist.");
                    return;
                }

                // Get worksheet dimensions
                var dimensions = worksheet.Dimension;

                // Define a dictionary to hold the headers and their positions
                var headers = new Dictionary<string, int>();

                // Loop through all rows in the first column and find the ones containing the headers we need
                for (var row = 1; row <= dimensions.End.Row; row++)
                {
                    var header = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                    if (header == null) continue;
                    headers.Add(header, row);
                }

                // Define a list to hold the XML keys and their replacements
                var replacements = new List<(string key, string value)>
                {
                    ("{SEAL_ID}", "SEAL ID"),
                    ("{TYPE_EN}", "TYPE"),
                    ("{TYPE_BG}", "ТИП"),
                    ("{GENERAL_LAYOUT_EN}", "GENERAL LAYOUT"),
                    ("{GENERAL_LAYOUT_BG}", "ОФОРМЛЕНИЕ"),
                    ("{MATRIX_EN}", "MATRIX"),
                    ("{MATRIX_BG}", "МАТРИЦА (ПЕЧАТ)"),
                    ("{TYPE_OF_IMPRESSION_EN}", "TYPE OF IMPRESSION"),
                    ("{TYPE_OF_IMPRESSION_BG}", "ОТПЕЧАТЪК"),
                    ("{MATERIAL_EN}", "MATERIAL"),
                    ("{MATERIAL_BG}", "МАТЕРИАЛ"),
                    ("{DIAMETER}", "DIMENSIONS (mm)"),
                    ("{WEIGHT}", "WEIGHT (g)"),
                    ("{AXIS}", "AXIS (clock)"),
                    ("{OVERSTRIKE_ORIENTATION}", "OVERSTRIKE ORIENTATION (clock)"),
                    ("{CHANNEL_ORIENTATION}", "CHANNEL ORIENTATION (clock)"),
                    ("{EXECUTION_EN}", "EXECUTION"),
                    ("{EXECUTION_BG}", "НАЧИН НА ИЗРАБОТВАНЕ"),
                    ("{COUNTERMARK_EN}", "COUNTERMARK"),
                    ("{COUNTERMARK_BG}", "КОНТРАМАРКИ"),
                    ("{LETTERING_EN}", "LETTERING"),
                    ("{LETTERING_BG}", "ОСОБЕНОСТИ НА БУКВИТЕ"),
                    ("{SHAPE_EN}", "SHAPE"),
                    ("{SHAPE_BG}", "ФОРМА НА ЯДРОТО"),
                    ("{CONDITION_EN}", "CONDITION"),
                    ("{CONDITION_BG}", "СЪВРЕМЕННО СЪСТОЯНИЕ"),
                    ("{DATE}", "DATE"),
                    ("{INTERNAL_DATE}", "INTERNAL DATE"),
                    ("{DATING_CRITERIA_EN}", "DATING CRITERIA"),
                    ("{DATING_CRITERIA_BG}", "КРИТЕРИИ ЗА ДАТИРАНЕ"),
                    ("{ALTERNATIVE_DATING_EN}", "ALTERNATIVE DATING"),
                    ("{ALTERNATIVE_DATING_BG}", "АЛТЕРНАТИВНА ДАТИРОВКА"),
                    ("{SEALS_CONTEXT_EN}", "SEAL’S CONTEXT"),
                    ("{SEALS_CONTEXT_BG}", "КОНТЕКСТ НА ПЕЧАТА"),
                    ("{ISSUER_EN}", "ISSUER"),
                    ("{ISSUER_BG}", "ИЗДАТЕЛ (СОБСТВЕНИК НА ПЕЧАТА)"),
                    ("{ISSUER_MILIEU_EN}","ISSUER’S MILIEU"),
                    ("{ISSUER_MILIEU_BG}","СФЕРА НА ДЕЙНОСТ НА ИЗДАТЕЛЯ (СОБСТВЕНИКА НА ПЕЧАТА)"),
                    ("{PLACE_OF_ORIGIN_EN}", "PLACE OF ORIGIN"),
                    ("{PLACE_OF_ORIGIN_BG}", "МЯСТО НА ИЗРАБОТКА"),
                    ("{FIND_PLACE_EN}", "FIND PLACE") ,
                    ("{FIND_PLACE_BG}", "МЕСТОНАМИРАНЕ"),
                    ("{FIND_DATE}", "FIND DATE"),
                    ("{FIND_CIRCUMSTANCES_EN}", "FIND CIRCUMSTANCES"),
                    ("{FIND_CIRCUMSTANCES_BG}", "ОБСТОЯТЕЛСТВА НА НАМИРАНЕ"),
                    ("{MODERN_LOCATION_EN}", "MODERN LOCATION"),
                    ("{MODERN_LOCATION_BG}", "СЪВРЕМЕННО СЕЛИЩЕ, ДО КОЕТО Е ОТКРИТ ПЕЧАТЪТ"),
                    ("{INSTITUTION_AND_REPOSITORY_EN}", "INSTITUTION AND REPOSITORY"),
                    ("{INSTITUTION_AND_REPOSITORY_BG}", "МЯСТО НА СЪХРАНЕНИЕ"),
                    ("{COLLECTION_AND_INVENTORY}", "COLLECTION AND INVENTORY"),
                    ("{ACQUISITION_EN}", "ACQUISITION"),
                    ("{ACQUISITION_BG}", "СПОСОБ НА ПРИДОБИВАНЕ"),
                    ("{PREVIOUS_LOCATIONS_EN}", "PREVIOUS LOCATIONS"),
                    ("{PREVIOUS_LOCATIONS_BG}", "ПРЕДИШНО МЕСТОСЪХРАНЕНИЕ"),
                    ("{MODERN_OBSERVATIONS_EN}", "MODERN OBSERVATIONS"),
                    ("{MODERN_OBSERVATIONS_BG}", "СЪВРЕМЕННИ НАБЛЮДЕНИЯ"),
                    ("{OBVERSE_LAYOUT_OF_FIELD_EN}",  "OBVERSE LAYOUT OF FIELD"),
                    ("{OBVERSE_LAYOUT_OF_FIELD_BG}", "ОФОРМЛЕНИЕ НА ЛИЦЕВАТА СТРАНА"),
                    ("{OBVERSE_FIELDS_DIMENSIONS}", "OBVERSE FIELD’S DIMENSIONS (mm)"),
                    ("{OBVERSE_MATRIX_EN}", "OBVERSE MATRIX"),
                    ("{OBVERSE_MATRIX_BG}", "ЛИЦЕВ ПЕЧАТ / ЛИЦЕВА МАТРИЦА"),
                    ("{OBVERSE_ICONOGRAPHY_EN}", "OBVERSE ICONOGRAPHY"),
                    ("{OBVERSE_ICONOGRAPHY_BG}", "ИКОНОГРАФИЯ НА АВЕРСА"),
                    ("{OBVERSE_DECORATION_EN}", "OBVERSE DECORATION"),
                    ("{OBVERSE_DECORATION_BG}", "ДЕКОРАТИВНИ ЕЛЕМЕНТИ НА АВЕРСА"),
                    ("{REVERSE_LAYOUT_FIELD_EN}", "REVERSE LAYOUT FIELD"),
                    ("{REVERSE_LAYOUT_FIELD_BG}", "ОФОРМЛЕНИЕ НА ОБРАТНАТА СТРАНА"),
                    ("{REVERSE_FIELDS_DIMENSIONS}", "REVERSE FIELD’S DIMENSIONS (mm)"),
                    ("{REVERSE_MATRIX_EN}", "REVERSE MATRIX"),
                    ("{REVERSE_MATRIX_BG}", "РЕВЕРСЕН ПЕЧАТ / РЕВЕРС НА МАТРИЦА"),
                    ("{REVERSE_ICONOGRAPHY_EN}", "REVERSE ICONOGRAPHY"),
                    ("{REVERSE_ICONOGRAPHY_BG}", "ИКОНОГРАФИЯ НА РЕВЕРСА"),
                    ("{REVERSE_DECORATION_EN}", "REVERSE DECORATION"),
                    ("{REVERSE_DECORATION_BG}", "ДЕКОРАТИВНИ ЕЛЕМЕНТИ НА РЕВЕРСА"),
                    ("{LANGUAGE_EN}", "LANGUAGE(S)"),
                    ("{LANGUAGE_BG}", "ЕЗИК (ЕЗИЦИ)"),
                    ("{EDITION}", "EDITION(S)"),
                    ("{COMMENTARY_ON_EDITION_EN}", "COMMENTARY ON EDITION(S)"),
                    ("{COMMENTARY_ON_EDITION_BG}", "КОМЕНТАР НА ПУБЛИКАЦИИТЕ"),
                    ("{PARALLEL_EN}", "PARALLEL(S)") ,
                    ("{PARALLEL_BG}", "ПАРАЛЕЛ (ПАРАЛЕЛИ)"),
                    ("{COMMENTARY_ON_PARALLEL_EN}", "COMMENTARY ON PARALLEL(S)"),
                    ("{COMMENTARY_ON_PARALLEL_BG}", "КОМЕНТАР НА ПАРАЛЕЛИТЕ"),
                    ("{EDITION_INTERPRETIVE_EN}", "EDITION INTERPRETIVE"),
                    ("{EDITION_INTERPRETIVE_BG}", "ИНТЕРПРЕТАТИВНО ИЗДАНИЕ"),
                    ("{EDITION_DIPLOMATIC_EN}", "EDITION DIPLOMATIC"),
                    ("{EDITION_DIPLOMATIC_BG}", "ДИПЛОМАТИЧНО ИЗДАНИЕ"),
                    ("{APPARATUS_EN}", "APPARATUS"),
                    ("{APPARATUS_BG}", "КРИТИЧЕН АПАРАТ"),
                    ("{LEGEND_EN}", "LEGEND"),
                    ("{LEGEND_BG}", "НАДПИСИ"),
                    ("{TRANSLATION_EN}", "TRANSLATION"),
                    ("{TRANSLATION_BG}", "ПРЕВОД НА НАДПИСИТЕ"),
                    ("{COMMENTARY_EN}", "COMMENTARY"),
                    ("{COMMENTARY_BG}", "КОМЕНТАР НА НАДПИСИТЕ"),
                    ("{FOOTNOTES_EN}", "FOOTNOTES"),
                    ("{FOOTNOTES_BG}", "БЕЛЕЖКИ ПОД ЛИНИЯ"),
                    ("{BIBLIOGRAPHY_EN}", "BIBLIOGRAPHY"),
                    ("{BIBLIOGRAPHY_BG}", "БИБЛИОГРАФИЯ"),
                    ("{TITLE_EN}", "TITLE"),
                    ("{TITLE_BG}", "ЗАГЛАВИЕ"),
                    ("{EDITOR_FORENAME_EN}", "EDITOR FORENAME"),
                    ("{EDITOR_FORENAME_BG}", "СОБСТВЕНО ИМЕ НА РЕДАКТОРА"),
                    ("{EDITOR_SURNAME_EN}", "EDITOR SURNAME"),
                    ("{EDITOR_SURNAME_BG}", "ФАМИЛНО ИМЕ НА РЕДАКТОРА"),
                    ("{FILENAME}", "FILENAME"),
                    ("{SEQUENCE}", "SEQUENCE"),
                    ("{NOT_BEFORE}", "NOT BEFORE"),
                    ("{NOT_AFTER}", "NOT AFTER"),
                    ("{}", "{}")
                };

                List<Coordinates>? coordinates = LoadCoordinatesFromFile($@"./resources/coordinates.json");

                // Loop through all subsequent columns and retrieve the data for each header
                for (var col = 2; col <= dimensions.End.Column; col++)
                {

                    // Define a dictionary to hold the header and its value
                    var allValues = new Dictionary<string, string>();

                    // Loop through all headers and get their values
                    foreach (var header in headers)
                    {
                        var value = worksheet.Cells[header.Value, col].Value?.ToString()?.Trim() ?? "-";
                        allValues.Add(header.Key, value);
                    }

                    // Generate a filename and sequence
                    var sealId = allValues["SEAL ID"];
                    var filename = $"TM_{sealId}";
                    var sequence = sealId.PadLeft(4, '0');

                    // Add the filename and sequence to the dictionary
                    allValues.Add("FILENAME", filename);
                    allValues.Add("SEQUENCE", sequence);

                    // Get not before and not after dates from internal date
                    var internalDate = allValues["INTERNAL DATE"];
                    var notBefore = internalDate.Split('-')[0].PadLeft(4, '0');
                    var notAfter = internalDate.Split('-')[1].PadLeft(4, '0');

                    // Add the not before and not after dates to the dictionary
                    allValues.Add("NOT BEFORE", notBefore);
                    allValues.Add("NOT AFTER", notAfter);

                    // Get coordinates for the current seal
                    var latitude = allValues["LATITUDE"];
                    var longitude = allValues["LONGITUDE"];

                    // Create new Coordinates object
                    var location = new Coordinates(sealId, latitude, longitude);

                    // Add coordinates to the list
                    coordinates?.Add(location);

                    // Add empty curly braces to the dictionary
                    allValues.Add("{}", "-");



                    // Replace the XML keys with the corresponding values
                    foreach (var element in xmlTemplate.Descendants())
                    {
                        foreach (var replacement in replacements)
                        {
                            if (element.Value == replacement.key)
                            {
                                element.Value = allValues[replacement.value];
                            }

                            foreach (var attribute in element.Attributes())
                            {
                                if (attribute.Value == replacement.key)
                                {
                                    attribute.Value = allValues[replacement.value];
                                }
                            }
                        }
                    }

                    // Save the updated XML file to disk
                    xmlTemplate.Save($"../webapps/ROOT/content/xml/epidoc/{filename}.xml");

                    // Reset the dictionary
                    allValues.Clear();

                    // Reset the XML document
                    xmlTemplate = XDocument.Load(@"./resources/SigiDocTemplate.xml");

                    // Get the <list> element from the seals list xml file
                    var listElement = sealsList.Descendants(ns + "list").FirstOrDefault();

                    if (listElement != null)
                    {
                        // Create a new list item
                        var newListItem = new XElement(ns + "item");
                        newListItem.SetAttributeValue("n", filename);
                        newListItem.SetAttributeValue("sortKey", sequence);
                        listElement.Add(newListItem);

                        // Save the updated seals list xml file to disk
                        sealsList.Save(@"../all_seals.xml");

                        // Reset the seals list xml document
                        sealsList = XDocument.Load(@"../all_seals.xml");
                    }
                }

                // Create serialization options with custom settings
                JsonSerializerOptions? options = new JsonSerializerOptions
                {
                    WriteIndented = true // Enable indented formatting
                };

                // Serialize the coordinates list to JSON
                var json = JsonSerializer.Serialize(coordinates, options);

                // Save the JSON to the file
                File.WriteAllText($@"./resources/coordinates.json", json);

                // Close the Excel file
                package.Dispose();

                Console.WriteLine("Success!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static List<Coordinates>? LoadCoordinatesFromFile(string path)
        {
            try
            {
                var json = File.ReadAllText(path);
                return JsonSerializer.Deserialize<List<Coordinates>>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return null;
            }
        }
    }

    internal class Coordinates
    {
        public Coordinates(string sealId, string latitude, string longitude)
        {
            SealId = sealId;
            Latitude = latitude;
            Longitude = longitude;
        }

        [JsonPropertyName("sealId")]
        public string SealId { get; set; }

        [JsonPropertyName("latitude")]
        public string Latitude { get; set; }

        [JsonPropertyName("longitude")]
        public string Longitude
        {
            get; set;
        }
    }
}