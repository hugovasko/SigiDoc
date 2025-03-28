using System.Collections.Generic;

namespace ExcelToXMLConverter
{
    internal static class ReplacementValues
    {
        public static List<(string key, string value)> Replacements { get; } = new List<(string key, string value)>
        {
            ("{SEAL_ID}", "SIGIDOC ID"),
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
            ("{DATE}", "ANALYSIS DATE"),
            ("{INTERNAL_DATE}", "INTERNAL DATE"),
            ("{ANALYSIS_DATE_CRITERIA_EN}", "ANALYSIS DATE CRITERIA"),
            ("{ANALYSIS_DATE_CRITERIA_BG}", "АНАЛИЗ НА ДАТИРОВКА – КРИТЕРИИ"),
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
            ("{FIND_PLACE_EN}", "FIND PLACE – ANCIENT FINDSPOT") ,
            ("{FIND_PLACE_BG}", "МЕСТОНАМИРАНЕ – АНТИЧЕН ТОПОНИМ"),
            ("{FIND_DATE}", "FIND DATE"),
            ("{FIND_CIRCUMSTANCES_EN}", "FIND CIRCUMSTANCES"),
            ("{FIND_CIRCUMSTANCES_BG}", "ОБСТОЯТЕЛСТВА НА НАМИРАНЕ"),
            ("{MODERN_LOCATION_EN}", "FIND PLACE – MODERN FINDSPOT"),
            ("{MODERN_LOCATION_BG}", "МЕСТОНАМИРАНЕ – СЪВРЕМЕНЕН ТОПОНИМ"),
            ("{INSTITUTION_EN}", "INSTITUTION"),
            ("{INSTITUTION_BG}", "ИНСТИТУЦИЯ"),
            ("{REPOSITORY_EN}", "REPOSITORY"),
            ("{REPOSITORY_BG}", "МЯСТО НА СЪХРАНЕНИЕ"),
            ("{COLLECTION_EN}", "COLLECTION"),
            ("{COLLECTION_BG}", "КОЛЕКЦИЯ"),
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
            ("{EDITION(S)_EN}", "EDITION(S)"),
            ("{EDITION(S)_BG}", "ПУБЛИКАЦИЯ (ПУБЛИКАЦИИ)"),
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
            ("{TITLE_EDITOR_FORENAME_EN}", "TITLE EDITOR FORENAME"),
            ("{TITLE_EDITOR_FORENAME_BG}", "РЕДАКТОР НА ЗАГЛАВИЕТО СОБСТВЕНО ИМЕ"),
            ("{TITLE_EDITOR_SURNAME_EN}", "TITLE EDITOR SURNAME"),
            ("{TITLE_EDITOR_SURNAME_BG}", "РЕДАКТОР НА ЗАГЛАВИЕТО ФАМИЛНО ИМЕ"),
            ("{EDITION_EDITOR_FORENAME_EN}", "EDITION EDITOR FORENAME"),
            ("{EDITION_EDITOR_FORENAME_BG}", "РЕДАКТОР НА ЗАПИСА СОБСТВЕНО ИМЕ"),
            ("{EDITION_EDITOR_SURNAME_EN}", "EDITION EDITOR SURNAME"),
            ("{EDITION_EDITOR_SURNAME_BG}", "РЕДАКТОР НА ЗАПИСА ФАМИЛНО ИМЕ"),
            ("{FILENAME}", "FILENAME"),
            ("{SEQUENCE}", "SEQUENCE"),
            ("{NOT_BEFORE}", "ANALYSIS DATE NOT BEFORE"),
            ("{NOT_AFTER}", "ANALYSIS DATE NOT AFTER"),
            ("{COORDINATES}", "COORDINATES"),
            ("{FINDSPOT_ACCURACY}", "FINDSPOT ACCURACY"),
            ("{FACSIMILE_OBVERSE_GRAPHIC}", "FACSIMILE OBVERSE GRAPHIC"),
            ("{FACSIMILE_REVERSE_GRAPHIC}", "FACSIMILE REVERSE GRAPHIC"),
            ("{FACSIMILE_OBVERSE_DESCRIPTION}", "FACSIMILE OBVERSE DESCRIPTION"),
            ("{FACSIMILE_REVERSE_DESCRIPTION}", "FACSIMILE REVERSE DESCRIPTION"),
            ("{}", "{}")
        };
    }
}