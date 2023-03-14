﻿using System.Xml.Linq;
using OfficeOpenXml;

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

                // Get rows headers from the first column (column A)
                var rowHeaders = worksheet.Cells[1, 1, dimensions.End.Row, 1]
                    .Select(c => c.Value?.ToString()?.Trim()).ToArray();

                // Find the rows containing the headers we need
                var sealIdRow = Array.IndexOf(rowHeaders, "SEAL ID") + 1;
                var typeEnRow = Array.IndexOf(rowHeaders, "TYPE") + 1;
                var typeBgRow = Array.IndexOf(rowHeaders, "ТИП") + 1;
                var generalLayoutEnRow = Array.IndexOf(rowHeaders, "GENERAL LAYOUT") + 1;
                var generalLayoutBgRow = Array.IndexOf(rowHeaders, "ОФОРМЛЕНИЕ") + 1;
                var matrixEnRow = Array.IndexOf(rowHeaders, "MATRIX") + 1;
                var matrixBgRow = Array.IndexOf(rowHeaders, "МАТРИЦА (ПЕЧАТ)") + 1;
                var typeOfImpressionEnRow = Array.IndexOf(rowHeaders, "TYPE OF IMPRESSION") + 1;
                var typeOfImpressionBgRow = Array.IndexOf(rowHeaders, "ОТПЕЧАТЪК") + 1;
                var materialEnRow = Array.IndexOf(rowHeaders, "MATERIAL") + 1;
                var materialBgRow = Array.IndexOf(rowHeaders, "МАТЕРИАЛ") + 1;
                var diameterRow = Array.IndexOf(rowHeaders, "DIMENSIONS (mm)") + 1;
                var weightRow = Array.IndexOf(rowHeaders, "WEIGHT (g)") + 1;
                var axisRow = Array.IndexOf(rowHeaders, "AXIS (clock)") + 1;
                var overstrikeOrientationRow = Array.IndexOf(rowHeaders, "OVERSTRIKE ORIENTATION (clock)") + 1;
                var channelOrientationRow = Array.IndexOf(rowHeaders, "CHANNEL ORIENTATION (clock)") + 1;
                var executionEnRow = Array.IndexOf(rowHeaders, "EXECUTION") + 1;
                var executionBgRow = Array.IndexOf(rowHeaders, "НАЧИН НА ИЗРАБОТВАНЕ") + 1;
                var countermarkEnRow = Array.IndexOf(rowHeaders, "COUNTERMARK") + 1;
                var countermarkBgRow = Array.IndexOf(rowHeaders, "КОНТРАМАРКИ") + 1;
                var letteringEnRow = Array.IndexOf(rowHeaders, "LETTERING") + 1;
                var letteringBgRow = Array.IndexOf(rowHeaders, "ОСОБЕНОСТИ НА БУКВИТЕ") + 1;
                var shapeEnRow = Array.IndexOf(rowHeaders, "SHAPE") + 1;
                var shapeBgRow = Array.IndexOf(rowHeaders, "ФОРМА НА ЯДРОТО") + 1;
                var conditionEnRow = Array.IndexOf(rowHeaders, "CONDITION") + 1;
                var conditionBgRow = Array.IndexOf(rowHeaders, "СЪВРЕМЕННО СЪСТОЯНИЕ") + 1;
                var dateRow = Array.IndexOf(rowHeaders, "DATE") + 1;
                var internalDateRow = Array.IndexOf(rowHeaders, "INTERNAL DATE") + 1;
                var datingCriteriaEnRow = Array.IndexOf(rowHeaders, "DATING CRITERIA") + 1;
                var datingCriteriaBgRow = Array.IndexOf(rowHeaders, "КРИТЕРИИ ЗА ДАТИРАНЕ") + 1;
                var alternativeDatingEnRow = Array.IndexOf(rowHeaders, "ALTERNATIVE DATING") + 1;
                var alternativeDatingBgRow = Array.IndexOf(rowHeaders, "АЛТЕРНАТИВНА ДАТИРОВКА") + 1;
                var sealsContextEnRow = Array.IndexOf(rowHeaders, "SEAL’S CONTEXT") + 1;
                var sealsContextBgRow = Array.IndexOf(rowHeaders, "КОНТЕКСТ НА ПЕЧАТА") + 1;
                var issuerEnRow = Array.IndexOf(rowHeaders, "ISSUER") + 1;
                var issuerBgRow = Array.IndexOf(rowHeaders, "ИЗДАТЕЛ (СОБСТВЕНИК НА ПЕЧАТА)") + 1;
                var issuersMilieuEnRow = Array.IndexOf(rowHeaders, "ISSUER’S MILIEU") + 1;
                var issuersMilieuBgRow = Array.IndexOf(rowHeaders, "СФЕРА НА ДЕЙНОСТ НА ИЗДАТЕЛЯ (СОБСТВЕНИКА НА ПЕЧАТА)") + 1;
                var placeOfOriginEnRow = Array.IndexOf(rowHeaders, "PLACE OF ORIGIN") + 1;
                var placeOfOriginBgRow = Array.IndexOf(rowHeaders, "МЯСТО НА ИЗРАБОТКА") + 1;
                var findPlaceEnRow = Array.IndexOf(rowHeaders, "FIND PLACE") + 1;
                var findPlaceBgRow = Array.IndexOf(rowHeaders, "МЕСТОНАМИРАНЕ") + 1;
                var findDateRow = Array.IndexOf(rowHeaders, "FIND DATE") + 1;
                var findCircumstancesEnRow = Array.IndexOf(rowHeaders, "FIND CIRCUMSTANCES") + 1;
                var findCircumstancesBgRow = Array.IndexOf(rowHeaders, "ОБСТОЯТЕЛСТВА НА НАМИРАНЕ") + 1;
                var modernLocationEnRow = Array.IndexOf(rowHeaders, "MODERN LOCATION") + 1;
                var modernLocationBgRow = Array.IndexOf(rowHeaders, "СЪВРЕМЕННО СЕЛИЩЕ, ДО КОЕТО Е ОТКРИТ ПЕЧАТЪТ") + 1;
                var institutionAndRepositoryEnRow = Array.IndexOf(rowHeaders, "INSTITUTION AND REPOSITORY") + 1;
                var institutionAndRepositoryBgRow = Array.IndexOf(rowHeaders, "МЯСТО НА СЪХРАНЕНИЕ ") + 1;
                var collectionAndInventoryRow = Array.IndexOf(rowHeaders, "COLLECTION AND INVENTORY") + 1;
                var acquisitionEnRow = Array.IndexOf(rowHeaders, "ACQUISITION") + 1;
                var acquisitionBgRow = Array.IndexOf(rowHeaders, "СПОСОБ НА ПРИДОБИВАНЕ") + 1;
                var previousLocationsEnRow = Array.IndexOf(rowHeaders, "PREVIOUS LOCATIONS") + 1;
                var previousLocationsBgRow = Array.IndexOf(rowHeaders, "ПРЕДИШНО МЕСТОСЪХРАНЕНИЕ") + 1;
                var modernObservationsEnRow = Array.IndexOf(rowHeaders, "MODERN OBSERVATIONS") + 1;
                var modernObservationsBgRow = Array.IndexOf(rowHeaders, "СЪВРЕМЕННИ НАБЛЮДЕНИЯ") + 1;
                var observeLayoutOfFieldEnRow = Array.IndexOf(rowHeaders, "OBSERVE LAYOUT OF FIELD") + 1;
                var observeLayoutOfFieldBgRow = Array.IndexOf(rowHeaders, "ОФОРМЛЕНИЕ НА ЛИЦЕВАТА СТРАНА") + 1;
                var observeFieldsDimensionsRow = Array.IndexOf(rowHeaders, "OBSERVE FIELD'S DIMENSIONS (mm)") + 1;
                var observeMatrixEnRow = Array.IndexOf(rowHeaders, "OBSERVE MATRIX") + 1;
                var observeMatrixBgRow = Array.IndexOf(rowHeaders, "ЛИЦЕВ ПЕЧАТ / ЛИЦЕВА МАТРИЦА") + 1;
                var observeIconographyEnRow = Array.IndexOf(rowHeaders, "OBSERVE ICONOGRAPHY") + 1;
                var observeIconographyBgRow = Array.IndexOf(rowHeaders, "ИКОНОГРАФИЯ НА АВЕРСА") + 1;
                var observeDecorationEnRow = Array.IndexOf(rowHeaders, "OBSERVE DECORATION") + 1;
                var observeDecorationBgRow = Array.IndexOf(rowHeaders, "ДЕКОРАТИВНИ ЕЛЕМЕНТИ НА АВЕРСА") + 1;
                var reverseLayoutFieldEnRow = Array.IndexOf(rowHeaders, "REVERSE LAYOUT FIELD") + 1;
                var reverseLayoutFieldBgRow = Array.IndexOf(rowHeaders, "ОФОРМЛЕНИЕ НА ОБРАТНАТА СТРАНА") + 1;
                var reverseFieldsDimensionsRow = Array.IndexOf(rowHeaders, "REVERSE FIELD'S DIMENSIONS (mm)") + 1;
                var reverseMatrixEnRow = Array.IndexOf(rowHeaders, "REVERSE MATRIX") + 1;
                var reverseMatrixBgRow = Array.IndexOf(rowHeaders, "РЕВЕРСЕН ПЕЧАТ / РЕВЕРС НА МАТРИЦА") + 1;
                var reverseIconographyEnRow = Array.IndexOf(rowHeaders, "REVERSE ICONOGRAPHY") + 1;
                var reverseIconographyBgRow = Array.IndexOf(rowHeaders, "ИКОНОГРАФИЯ НА РЕВЕРСА") + 1;
                var reverseDecorationEnRow = Array.IndexOf(rowHeaders, "REVERSE DECORATION") + 1;
                var reverseDecorationBgRow = Array.IndexOf(rowHeaders, "ДЕКОРАТИВНИ ЕЛЕМЕНТИ НА РЕВЕРСА") + 1;
                var languageEnRow = Array.IndexOf(rowHeaders, "LANGUAGE(S)") + 1;
                var languageBgRow = Array.IndexOf(rowHeaders, "ЕЗИК (ЕЗИЦИ)") + 1;
                var editionRow = Array.IndexOf(rowHeaders, "EDITION(S)") + 1;
                var commentaryOnEditionEnRow = Array.IndexOf(rowHeaders, "COMMENTARY ON EDITION(S)") + 1;
                var commentaryOnEditionBgRow = Array.IndexOf(rowHeaders, "КОМЕНТАР НА ПУБЛИКАЦИИТЕ") + 1;
                var parallelEnRow = Array.IndexOf(rowHeaders, "PARALLEL(S)") + 1;
                var parallelBgRow = Array.IndexOf(rowHeaders, "ПАРАЛЕЛ (ПАРАЛЕЛИ)") + 1;
                var commentaryOnParallelEnRow = Array.IndexOf(rowHeaders, "COMMENTARY ON PARALLEL(S)") + 1;
                var commentaryOnParallelBgRow = Array.IndexOf(rowHeaders, "КОМЕНТАР НА ПАРАЛЕЛИТЕ") + 1;
                var editionInterpretiveEnRow = Array.IndexOf(rowHeaders, "EDITION INTERPRETIVE") + 1;
                var editionInterpretiveBgRow = Array.IndexOf(rowHeaders, "ИНТЕРПРЕТАТИВНО ИЗДАНИЕ") + 1;
                var editionDiplomaticEnRow = Array.IndexOf(rowHeaders, "EDITION DIPLOMATIC") + 1;
                var editionDiplomaticBgRow = Array.IndexOf(rowHeaders, "ДИПЛОМАТИЧНО ИЗДАНИЕ") + 1;
                var apparatusEnRow = Array.IndexOf(rowHeaders, "APPARATUS") + 1;
                var apparatusBgRow = Array.IndexOf(rowHeaders, "КРИТИЧЕН АПАРАТ") + 1;
                var legendEnRow = Array.IndexOf(rowHeaders, "LEGEND") + 1;
                var legendBgRow = Array.IndexOf(rowHeaders, "НАДПИСИ") + 1;
                var translationEnRow = Array.IndexOf(rowHeaders, "TRANSLATION") + 1;
                var translationBgRow = Array.IndexOf(rowHeaders, "ПРЕВОД НА НАДПИСИТЕ") + 1;
                var commentaryEnRow = Array.IndexOf(rowHeaders, "COMMENTARY") + 1;
                var commentaryBgRow = Array.IndexOf(rowHeaders, "КОМЕНТАР НА НАДПИСИТЕ") + 1;
                var footnotesEnRow = Array.IndexOf(rowHeaders, "FOOTNOTES") + 1;
                var footnotesBgRow = Array.IndexOf(rowHeaders, "БЕЛЕЖКИ ПОД ЛИНИЯ") + 1;
                var bibliographyEnRow = Array.IndexOf(rowHeaders, "BIBLIOGRAPHY") + 1;
                var bibliographyBgRow = Array.IndexOf(rowHeaders, "БИБЛИОГРАФИЯ") + 1;
                var titleEnRow = Array.IndexOf(rowHeaders, "TITLE") + 1;
                var titleBgRow = Array.IndexOf(rowHeaders, "ЗАГЛАВИЕ") + 1;
                var editorForenameEnRow = Array.IndexOf(rowHeaders, "EDITOR FORENAME") + 1;
                var editorForenameBgRow = Array.IndexOf(rowHeaders, "СОБСТВЕНО ИМЕ НА РЕДАКТОРА") + 1;
                var editorSurnameEnRow = Array.IndexOf(rowHeaders, "EDITOR SURNAME") + 1;
                var editorSurnameBgRow = Array.IndexOf(rowHeaders, "ФАМИЛНО ИМЕ НА РЕДАКТОРА") + 1;

                // Loop through all subsequent columns and retrieve the data for each header
                for (var col = 2; col <= dimensions.End.Column; col++)
                {
                    // Get the values for each header
                    var sealId = worksheet.Cells[sealIdRow, col].Value?.ToString() ?? "-";
                    var typeEn = worksheet.Cells[typeEnRow, col].Value?.ToString() ?? "-";
                    var typeBg = worksheet.Cells[typeBgRow, col].Value?.ToString() ?? "-";
                    var generalLayoutEn = worksheet.Cells[generalLayoutEnRow, col].Value?.ToString() ?? "-";
                    var generalLayoutBg = worksheet.Cells[generalLayoutBgRow, col].Value?.ToString() ?? "-";
                    var matrixEn = worksheet.Cells[matrixEnRow, col].Value?.ToString() ?? "-";
                    var matrixBg = worksheet.Cells[matrixBgRow, col].Value?.ToString() ?? "-";
                    var typeOfImpressionEn = worksheet.Cells[typeOfImpressionEnRow, col].Value?.ToString() ?? "-";
                    var typeOfImpressionBg = worksheet.Cells[typeOfImpressionBgRow, col].Value?.ToString() ?? "-";
                    var materialEn = worksheet.Cells[materialEnRow, col].Value?.ToString() ?? "-";
                    var materialBg = worksheet.Cells[materialBgRow, col].Value?.ToString() ?? "-";
                    var diameter = worksheet.Cells[diameterRow, col].Value?.ToString() ?? "-";
                    var weight = worksheet.Cells[weightRow, col].Value?.ToString() ?? "-";
                    var axis = worksheet.Cells[axisRow, col].Value?.ToString() ?? "-";
                    var overstrikeOrientation = worksheet.Cells[overstrikeOrientationRow, col].Value?.ToString() ?? "-";
                    var channelOrientation = worksheet.Cells[channelOrientationRow, col].Value?.ToString() ?? "-";
                    var executionEn = worksheet.Cells[executionEnRow, col].Value?.ToString() ?? "-";
                    var executionBg = worksheet.Cells[executionBgRow, col].Value?.ToString() ?? "-";
                    var countermarkEn = worksheet.Cells[countermarkEnRow, col].Value?.ToString() ?? "-";
                    var countermarkBg = worksheet.Cells[countermarkBgRow, col].Value?.ToString() ?? "-";
                    var letteringEn = worksheet.Cells[letteringEnRow, col].Value?.ToString() ?? "-";
                    var letteringBg = worksheet.Cells[letteringBgRow, col].Value?.ToString() ?? "-";
                    var shapeEn = worksheet.Cells[shapeEnRow, col].Value?.ToString() ?? "-";
                    var shapeBg = worksheet.Cells[shapeBgRow, col].Value?.ToString() ?? "-";
                    var conditionEn = worksheet.Cells[conditionEnRow, col].Value?.ToString() ?? "-";
                    var conditionBg = worksheet.Cells[conditionBgRow, col].Value?.ToString() ?? "-";
                    var date = worksheet.Cells[dateRow, col].Value?.ToString() ?? "-";
                    var internalDate = worksheet.Cells[internalDateRow, col].Value?.ToString() ?? "-";
                    var datingCriteriaEn = worksheet.Cells[datingCriteriaEnRow, col].Value?.ToString() ?? "-";
                    var datingCriteriaBg = worksheet.Cells[datingCriteriaBgRow, col].Value?.ToString() ?? "-";
                    var alternativeDatingEn = worksheet.Cells[alternativeDatingEnRow, col].Value?.ToString() ?? "-";
                    var alternativeDatingBg = worksheet.Cells[alternativeDatingBgRow, col].Value?.ToString() ?? "-";
                    var sealsContextEn = worksheet.Cells[sealsContextEnRow, col].Value?.ToString() ?? "-";
                    var sealsContextBg = worksheet.Cells[sealsContextBgRow, col].Value?.ToString() ?? "-";
                    var issuerEn = worksheet.Cells[issuerEnRow, col].Value?.ToString() ?? "-";
                    var issuerBg = worksheet.Cells[issuerBgRow, col].Value?.ToString() ?? "-";
                    var issuersMilieuEn = worksheet.Cells[issuersMilieuEnRow, col].Value?.ToString() ?? "-";
                    var issuersMilieuBg = worksheet.Cells[issuersMilieuBgRow, col].Value?.ToString() ?? "-";
                    var placeOfOriginEn = worksheet.Cells[placeOfOriginEnRow, col].Value?.ToString() ?? "-";
                    var placeOfOriginBg = worksheet.Cells[placeOfOriginBgRow, col].Value?.ToString() ?? "-";
                    var findPlaceEn = worksheet.Cells[findPlaceEnRow, col].Value?.ToString() ?? "-";
                    var findPlaceBg = worksheet.Cells[findPlaceBgRow, col].Value?.ToString() ?? "-";
                    var findDate = worksheet.Cells[findDateRow, col].Value?.ToString() ?? "-";
                    var findCircumstancesEn = worksheet.Cells[findCircumstancesEnRow, col].Value?.ToString() ?? "-";
                    var findCircumstancesBg = worksheet.Cells[findCircumstancesBgRow, col].Value?.ToString() ?? "-";
                    var modernLocationEn = worksheet.Cells[modernLocationEnRow, col].Value?.ToString() ?? "-";
                    var modernLocationBg = worksheet.Cells[modernLocationBgRow, col].Value?.ToString() ?? "-";
                    var institutionAndRepositoryEn = worksheet.Cells[institutionAndRepositoryEnRow, col].Value?.ToString() ?? "-";
                    var institutionAndRepositoryBg = worksheet.Cells[institutionAndRepositoryBgRow, col].Value?.ToString() ?? "-";
                    var collectionAndInventory = worksheet.Cells[collectionAndInventoryRow, col].Value?.ToString() ?? "-";
                    var acquisitionEn = worksheet.Cells[acquisitionEnRow, col].Value?.ToString() ?? "-";
                    var acquisitionBg = worksheet.Cells[acquisitionBgRow, col].Value?.ToString() ?? "-";
                    var previousLocationsEn = worksheet.Cells[previousLocationsEnRow, col].Value?.ToString() ?? "-";
                    var previousLocationsBg = worksheet.Cells[previousLocationsBgRow, col].Value?.ToString() ?? "-";
                    var modernObservationsEn = worksheet.Cells[modernObservationsEnRow, col].Value?.ToString() ?? "-";
                    var modernObservationsBg = worksheet.Cells[modernObservationsBgRow, col].Value?.ToString() ?? "-";
                    var observeLayoutOfFieldEn = worksheet.Cells[observeLayoutOfFieldEnRow, col].Value?.ToString() ?? "-";
                    var observeLayoutOfFieldBg = worksheet.Cells[observeLayoutOfFieldBgRow, col].Value?.ToString() ?? "-";
                    var observeFieldsDimensions = worksheet.Cells[observeFieldsDimensionsRow, col].Value?.ToString() ?? "-";
                    var observeMatrixEn = worksheet.Cells[observeMatrixEnRow, col].Value?.ToString() ?? "-";
                    var observeMatrixBg = worksheet.Cells[observeMatrixBgRow, col].Value?.ToString() ?? "-";
                    var observeIconographyEn = worksheet.Cells[observeIconographyEnRow, col].Value?.ToString() ?? "-";
                    var observeIconographyBg = worksheet.Cells[observeIconographyBgRow, col].Value?.ToString() ?? "-";
                    var observeDecorationEn = worksheet.Cells[observeDecorationEnRow, col].Value?.ToString() ?? "-";
                    var observeDecorationBg = worksheet.Cells[observeDecorationBgRow, col].Value?.ToString() ?? "-";
                    var reverseLayoutFieldEn = worksheet.Cells[reverseLayoutFieldEnRow, col].Value?.ToString() ?? "-";
                    var reverseLayoutFieldBg = worksheet.Cells[reverseLayoutFieldBgRow, col].Value?.ToString() ?? "-";
                    var reverseFieldsDimensions = worksheet.Cells[reverseFieldsDimensionsRow, col].Value?.ToString() ?? "-";
                    var reverseMatrixEn = worksheet.Cells[reverseMatrixEnRow, col].Value?.ToString() ?? "-";
                    var reverseMatrixBg = worksheet.Cells[reverseMatrixBgRow, col].Value?.ToString() ?? "-";
                    var reverseIconographyEn = worksheet.Cells[reverseIconographyEnRow, col].Value?.ToString() ?? "-";
                    var reverseIconographyBg = worksheet.Cells[reverseIconographyBgRow, col].Value?.ToString() ?? "-";
                    var reverseDecorationEn = worksheet.Cells[reverseDecorationEnRow, col].Value?.ToString() ?? "-";
                    var reverseDecorationBg = worksheet.Cells[reverseDecorationBgRow, col].Value?.ToString() ?? "-";
                    var languageEn = worksheet.Cells[languageEnRow, col].Value?.ToString() ?? "-";
                    var languageBg = worksheet.Cells[languageBgRow, col].Value?.ToString() ?? "-";
                    var edition = worksheet.Cells[editionRow, col].Value?.ToString() ?? "-";
                    var commentaryOnEditionEn = worksheet.Cells[commentaryOnEditionEnRow, col].Value?.ToString() ?? "-";
                    var commentaryOnEditionBg = worksheet.Cells[commentaryOnEditionBgRow, col].Value?.ToString() ?? "-";
                    var parallelEn = worksheet.Cells[parallelEnRow, col].Value?.ToString() ?? "-";
                    var parallelBg = worksheet.Cells[parallelBgRow, col].Value?.ToString() ?? "-";
                    var commentaryOnParallelEn = worksheet.Cells[commentaryOnParallelEnRow, col].Value?.ToString() ?? "-";
                    var commentaryOnParallelBg = worksheet.Cells[commentaryOnParallelBgRow, col].Value?.ToString() ?? "-";
                    var editionInterpretiveEn = worksheet.Cells[editionInterpretiveEnRow, col].Value?.ToString() ?? "-";
                    var editionInterpretiveBg = worksheet.Cells[editionInterpretiveBgRow, col].Value?.ToString() ?? "-";
                    var editionDiplomaticEn = worksheet.Cells[editionDiplomaticEnRow, col].Value?.ToString() ?? "-";
                    var editionDiplomaticBg = worksheet.Cells[editionDiplomaticBgRow, col].Value?.ToString() ?? "-";
                    var apparatusEn = worksheet.Cells[apparatusEnRow, col].Value?.ToString() ?? "-";
                    var apparatusBg = worksheet.Cells[apparatusBgRow, col].Value?.ToString() ?? "-";
                    var legendEn = worksheet.Cells[legendEnRow, col].Value?.ToString() ?? "-";
                    var legendBg = worksheet.Cells[legendBgRow, col].Value?.ToString() ?? "-";
                    var translationEn = worksheet.Cells[translationEnRow, col].Value?.ToString() ?? "-";
                    var translationBg = worksheet.Cells[translationBgRow, col].Value?.ToString() ?? "-";
                    var commentaryEn = worksheet.Cells[commentaryEnRow, col].Value?.ToString() ?? "-";
                    var commentaryBg = worksheet.Cells[commentaryBgRow, col].Value?.ToString() ?? "-";
                    var footnotesEn = worksheet.Cells[footnotesEnRow, col].Value?.ToString() ?? "-";
                    var footnotesBg = worksheet.Cells[footnotesBgRow, col].Value?.ToString() ?? "-";
                    var bibliographyEn = worksheet.Cells[bibliographyEnRow, col].Value?.ToString() ?? "-";
                    var bibliographyBg = worksheet.Cells[bibliographyBgRow, col].Value?.ToString() ?? "-";
                    var titleEn = worksheet.Cells[titleEnRow, col].Value?.ToString() ?? "-";
                    var titleBg = worksheet.Cells[titleBgRow, col].Value?.ToString() ?? "-";
                    var editorForenameEn = worksheet.Cells[editorForenameEnRow, col].Value?.ToString() ?? "-";
                    var editorForenameBg = worksheet.Cells[editorForenameBgRow, col].Value?.ToString() ?? "-";
                    var editorSurnameEn = worksheet.Cells[editorSurnameEnRow, col].Value?.ToString() ?? "-";
                    var editorSurnameBg = worksheet.Cells[editorSurnameBgRow, col].Value?.ToString() ?? "-";

                    // Generate filename
                    var filename = $"TM_{sealId}";

                    // Generate sequence
                    var sequence = sealId.PadLeft(4, '0');

                    // Get not before and not after dates from internal date
                    var notBefore = internalDate.Split('-')[0].PadLeft(4, '0');
                    var notAfter = internalDate.Split('-')[1].PadLeft(4, '0');

                    // Define a dictionary that maps the keys to the corresponding values
                    var allValues = new Dictionary<string, string>
                    {
                        {"{SEAL_ID}", sealId},
                        {"{TITLE_EN}", titleEn},
                        {"{TITLE_BG}", titleBg},
                        {"{GENERAL_LAYOUT_EN}", generalLayoutEn},
                        {"{GENERAL_LAYOUT_BG}", generalLayoutBg},
                        {"{MATRIX_EN}", matrixEn},
                        {"{MATRIX_BG}", matrixBg},
                        {"{TYPE_OF_IMPRESSION_EN}", typeOfImpressionEn},
                        {"{TYPE_OF_IMPRESSION_BG}", typeOfImpressionBg},
                        {"{MATERIAL_EN}", materialEn},
                        {"{MATERIAL_BG}", materialBg},
                        {"{DIAMETER}", diameter},
                        {"{WEIGHT}", weight},
                        {"{AXIS}", axis},
                        {"{OVERSTRIKE_ORIENTATION}", overstrikeOrientation},
                        {"{CHANNEL_ORIENTATION}", channelOrientation},
                        {"{EXECUTION_EN}", executionEn},
                        {"{EXECUTION_BG}", executionBg}

                    };

                    // Replace the XML keys with the corresponding values
                    foreach (var element in xmlTemplate.Descendants())
                    {
                        if (allValues.TryGetValue(element.Value, out var elementReplacement))
                        {
                            element.Value = elementReplacement;
                        }
                        foreach (var attribute in element.Attributes())
                        {
                            if (allValues.TryGetValue(attribute.Value, out var attributeReplacement))
                            {
                                attribute.Value = attributeReplacement;
                            }
                        }
                    }

                    // Save the updated XML file to disk
                    xmlTemplate.Save($"./resources/{filename}.xml");


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

                // Close the Excel file
                package.Dispose();

                Console.WriteLine("Success!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}