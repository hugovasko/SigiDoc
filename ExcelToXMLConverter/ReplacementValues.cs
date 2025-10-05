using System.Collections.Generic;

namespace ExcelToXMLConverter
{
    internal static class ReplacementValues
    {
        public static List<(string key, string value)> Replacements { get; } = new List<(string key, string value)>
        {
            // IDs and Metadata
            ("{SEAL_ID}", "SIGIDOC ID"),
            ("{FILENAME}", "FILENAME"),
            ("{SEQUENCE}", "SEQUENCE"),
            ("{DATE}", "ANALYSIS DATE"),
            ("{NOT_BEFORE}", "ANALYSIS DATE NOT BEFORE"),
            ("{NOT_AFTER}", "ANALYSIS DATE NOT AFTER"),
            
            // Title and Editors
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
            
            // Authority (using existing fields or placeholder)
            ("{AUTHORITY_EN}", "ARTEFACT"),
            ("{AUTHORITY_BG}", "ПРЕДМЕТ"),
            
            // Type and Layout
            ("{TYPE_EN}", "TYPE"),
            ("{TYPE_BG}", "ТИП"),
            ("{GENERAL_LAYOUT_EN}", "GENERAL LAYOUT"),
            ("{GENERAL_LAYOUT_BG}", "ОФОРМЛЕНИЕ"),
            ("{MATRIX_EN}", "MATRIX"),
            ("{MATRIX_BG}", "МАТРИЦА (ПЕЧАТ)"),
            ("{TYPE_OF_IMPRESSION_EN}", "TYPE OF IMPRESSION"),
            ("{TYPE_OF_IMPRESSION_BG}", "ОТПЕЧАТЪК"),
            
            // Material and Physical Properties
            ("{MATERIAL_EN}", "MATERIAL"),
            ("{MATERIAL_BG}", "МАТЕРИАЛ"),
            ("{SHAPE_EN}", "SHAPE"),
            ("{SHAPE_BG}", "ФОРМА НА ЯДРОТО"),
            ("{DIAMETER}", "DIMENSIONS (mm)"),
            ("{THICKNESS}", "THICKNESS (mm)"),
            ("{HEIGHT}", "HEIGHT"),  // Placeholder - will be "―" if not in Excel
            ("{WIDTH}", "WIDTH"),    // Placeholder - will be "―" if not in Excel
            ("{DEPTH}", "DEPTH"),    // Placeholder - will be "―" if not in Excel
            ("{WEIGHT}", "WEIGHT (g)"),
            ("{AXIS}", "AXIS (CLOCK)"),
            ("{OVERSTRIKE_ORIENTATION}", "OVERSTRIKE ORIENTATION (CLOCK)"),
            ("{CHANNEL_ORIENTATION}", "CHANNEL ORIENTATION (CLOCK)"),
            ("{EXECUTION_EN}", "EXECUTION"),
            ("{EXECUTION_BG}", "НАЧИН НА ИЗРАБОТВАНЕ"),
            ("{COUNTERMARK_EN}", "COUNTERMARK"),
            ("{COUNTERMARK_BG}", "КОНТРАМАРКИ"),
            ("{CONDITION_EN}", "CONDITION"),
            ("{CONDITION_BG}", "СЪВРЕМЕННО СЪСТОЯНИЕ"),
            
            // Whole document hand notes and decoration
            ("{HAND_NOTE_EN}", "LETTERING"),
            ("{HAND_NOTE_BG}", "ОСОБЕНОСТИ НА БУКВИТЕ"),
            ("{DECO_NOTE_EN}", "PHYSICAL DESCRIPTION"),
            ("{DECO_NOTE_BG}", "ОПИСАНИЕ"),
            
            // Dating
            ("{INTERNAL_DATE}", "INTERNAL DATE"),
            ("{INTERNAL_DATE_CRITERIA_EN}", "INTERNAL DATE CRITERIA"),
            ("{INTERNAL_DATE_CRITERIA_BG}", "КРИТЕРИИ НА ДАТИРОВКАТА, УКАЗАНА В ПЕЧАТА"),
            ("{ANALYSIS_DATE_CRITERIA_EN}", "ANALYSIS DATE CRITERIA"),
            ("{ANALYSIS_DATE_CRITERIA_BG}", "АНАЛИЗ НА ДАТИРОВКА – КРИТЕРИИ"),
            ("{ALTERNATIVE_DATING_EN}", "ALTERNATIVE DATING"),
            ("{ALTERNATIVE_DATING_BG}", "АЛТЕРНАТИВНА ДАТИРОВКА"),
            
            // Context and Issuer
            ("{SUMMARY_EN}", "CATEGORY"),
            ("{SUMMARY_BG}", "КАТЕГОРИЯ"),
            ("{SEALS_CONTEXT_EN}", "SEAL'S CONTEXT"),
            ("{SEALS_CONTEXT_BG}", "КОНТЕКСТ НА ПЕЧАТА"),
            ("{ISSUER_EN}", "ISSUER"),
            ("{ISSUER_BG}", "ИЗДАТЕЛ (СОБСТВЕНИК НА ПЕЧАТА)"),
            ("{ISSUER_MILIEU_EN}", "ISSUER'S MILIEU"),
            ("{ISSUER_MILIEU_BG}", "СФЕРА НА ДЕЙНОСТ НА ИЗДАТЕЛЯ (СОБСТВЕНИКА НА ПЕЧАТА)"),
            
            // Location - Origin and Finding
            ("{PLACE_OF_ORIGIN_EN}", "PLACE OF ORIGIN"),
            ("{PLACE_OF_ORIGIN_BG}", "МЯСТО НА ИЗРАБОТКА"),
            ("{FIND_PLACE_EN}", "FIND PLACE – ANCIENT FINDSPOT"),
            ("{FIND_PLACE_BG}", "МЕСТОНАМИРАНЕ – АНТИЧЕН ТОПОНИМ"),
            ("{MODERN_LOCATION_EN}", "FIND PLACE _ MODERN FINDSPOT"),
            ("{MODERN_LOCATION_BG}", "МЕСТОНАМИРАНЕ – СЪВРЕМЕНЕН ТОПОНИМ"),
            ("{FIND_DATE}", "FIND DATE"),
            ("{FIND_CIRCUMSTANCES_EN}", "FIND CIRCUMSTANCES"),
            ("{FIND_CIRCUMSTANCES_BG}", "ОБСТОЯТЕЛСТВА НА НАМИРАНЕ"),
            ("{COORDINATES}", "COORDINATES"),
            ("{FINDSPOT_ACCURACY}", "FINDSPOT ACCURACY"),
            
            // Modern Location and Repository
            ("{COUNTRY_EN}", "COUNTRY"),
            ("{COUNTRY_BG}", "ДЪРЖАВА"),
            ("{SETTLEMENT_EN}", "SETTLEMENT"),
            ("{SETTLEMENT_BG}", "СЕЛИЩЕ"),
            ("{INSTITUTION_EN}", "INSTITUTION"),
            ("{INSTITUTION_BG}", "ИНСТИТУЦИЯ"),
            ("{REPOSITORY_EN}", "REPOSITORY"),
            ("{REPOSITORY_BG}", "МЯСТО НА СЪХРАНЕНИЕ"),
            ("{COLLECTION_EN}", "COLLECTION"),
            ("{COLLECTION_BG}", "КОЛЕКЦИЯ"),
            ("{IDNO}", "INVENTORY NUMBER"),
            ("{ACQUISITION_EN}", "ACQUISITION"),
            ("{ACQUISITION_BG}", "СПОСОБ НА ПРИДОБИВАНЕ"),
            ("{PREVIOUS_LOCATIONS_EN}", "PREVIOUS LOCATIONS"),
            ("{PREVIOUS_LOCATIONS_BG}", "ПРЕДИШНО МЕСТОСЪХРАНЕНИЕ"),
            ("{MODERN_OBSERVATIONS_EN}", "MODERN OBSERVATIONS"),
            ("{MODERN_OBSERVATIONS_BG}", "СЪВРЕМЕНИ НАБЛЮДЕНИЯ"),
            
            // Obverse (r) fields
            ("{MS_CONTENTS_SUMMARY_R_EN}", "OBVERSE"),
            ("{MS_CONTENTS_SUMMARY_R_BG}", "АВЕРС"),
            ("{LAYOUT_R_EN}", "OBVERSE LAYOUT OF FIELD"),
            ("{LAYOUT_R_BG}", "ОФОРМЛЕНИЕ НА ЛИЦЕВАТА СТРАНА"),
            ("{DIAMETER_R}", "OBVERSE FIELD'S DIMENSIONS (mm)"),
            ("{MATRIX_R_EN}", "OBVERSE MATRIX"),
            ("{MATRIX_R_BG}", "ЛИЦЕВ ПЕЧАТ / ЛИЦЕВА МАТРИЦА"),
            ("{FIGURE_R_EN}", "OBVERSE ICONOGRAPHY"),
            ("{FIGURE_R_BG}", "ИКОНОГРАФИЯ НА АВЕРСА"),
            ("{FIGURE_DECOR_EN}", "OBVERSE DECORATION"),
            ("{FIGURE_DECOR_BG}", "ДЕКОРАТИВНИ ЕЛЕМЕНТИ НА АВЕРСА"),
            ("{HAND_NOTE_R_EN}", "OBVERSE EPIGRAPHY"),
            ("{HAND_NOTE_R_BG}", "ЕПИГРАФИКА НА АВЕРСА"),
            ("{TEXTLANG_R_EN}", "OBVERSE LANGUAGE"),
            ("{TEXTLANG_R_BG}", "ЕЗИК (ЕЗИЦИ) НА АВЕРСА"),
            
            // Reverse (v) fields
            ("{MS_CONTENTS_SUMMARY_V_EN}", "REVERSE"),
            ("{MS_CONTENTS_SUMMARY_V_BG}", "РЕВЕРС"),
            ("{LAYOUT_V_EN}", "REVERSE LAYOUT OF FIELD"),
            ("{LAYOUT_V_BG}", "ОФОРМЛЕНИЕ НА ОБРАТНАТА СТРАНА"),
            ("{DIAMETER_V}", "REVERSE FIELD'S DIMENSIONS (mm)"),
            ("{MATRIX_V_EN}", "REVERSE MATRIX"),
            ("{MATRIX_V_BG}", "РЕВЕРСЕН ПЕЧАТ / РЕВЕРС НА МАТРИЦА"),
            ("{FIGURE_V_EN}", "REVERSE ICONOGRAPHY"),
            ("{FIGURE_V_BG}", "ИКОНОГРАФИЯ НА РЕВЕРСА"),
            ("{FIGURE_DECOV_EN}", "REVERSE DECORATION"),
            ("{FIGURE_DECOV_BG}", "ДЕКОРАТОВНИ ЕЛЕМЕНТИ НА РЕВЕРСА"),
            ("{HAND_NOTE_V_EN}", "REVERSE EPIGRAPHY"),
            ("{HAND_NOTE_V_BG}", "ЕПИГРАФИКА НА РЕВЕРСА"),
            ("{TEXTLANG_V_EN}", "REVERSE LANGUAGE"),
            ("{TEXTLANG_V_BG}", "ЕЗИК (ЕЗИЦИ) НА РЕВЕРСА"),
            
            // Facsimile
            ("{FACSIMILE_OBVERSE_GRAPHIC}", "FACSIMILE OBVERSE GRAPHIC"),
            ("{FACSIMILE_REVERSE_GRAPHIC}", "FACSIMILE REVERSE GRAPHIC"),
            ("{FACSIMILE_OBVERSE_DESCRIPTION}", "FACSIMILE OBVERSE DESCRIPTION"),
            ("{FACSIMILE_REVERSE_DESCRIPTION}", "FACSIMILE REVERSE DESCRIPTION"),
            
            // Edition and Text
            ("{EDITION(S)_EN}", "EDITION(S)"),
            ("{EDITION(S)_BG}", "ПУБЛИКАЦИЯ (ПУБЛИКАЦИИ)"),
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
            ("{COMMENTARY_EN}", "COMMENTARY ON TEXT"),
            ("{COMMENTARY_BG}", "КОМЕНТАР НА НАДПИСИТЕ"),
            
            // Bibliography
            ("{PARALLEL_EN}", "PARALLEL(S)"),
            ("{PARALLEL_BG}", "ПАРАЛЕЛ (ПАРАЛЕЛИ)"),
            ("{COMMENTARY_ON_EDITION_EN}", "COMMENTARY ON EDITION(S)"),
            ("{COMMENTARY_ON_EDITION_BG}", "КОМЕНТАР НА ПУБЛИКАЦИИТЕ"),
            ("{COMMENTARY_ON_PARALLEL_EN}", "COMMENTARY ON PARALLEL(S)"),
            ("{COMMENTARY_ON_PARALLEL_BG}", "КОМЕНТАР НА ПАРАЛЕЛИТЕ"),
            ("{FOOTNOTES_EN}", "FOOTNOTES"),
            ("{FOOTNOTES_BG}", "БЕЛЕЖКИ ПОД ЛИНИЯ"),
            ("{BIBLIOGRAPHY_EN}", "BIBLIOGRAPHY"),
            ("{BIBLIOGRAPHY_BG}", "БИБЛИОГРАФИЯ"),

            // Apparatus
            ("{APPARATUS_EN}", "APPARATUS"),
            ("{APPARATUS_BG}", "КРИТИЧЕН АПАРАТ"),

            // Legend
            ("{LEGEND_EN}", "LEGEND"),
            ("{LEGEND_BG}", "НАДПИСИ"),

            // Translation
            ("{TRANSLATION_EN}", "TRANSLATION"),
            ("{TRANSLATION_BG}", "ПРЕВОД НА НАДПИСИТЕ"),

            // Commentary
            ("{COMMENTARY_TEXT_EN}", "COMMENTARY ON TEXT"),
            ("{COMMENTARY_TEXT_BG}", "КОМЕНТАР НА НАДПИСИТЕ"),

            // Footnotes
            ("{FOOTNOTES_EN}", "FOOTNOTES"),
            ("{FOOTNOTES_BG}", "БЕЛЕЖКИ ПОД ЛИНИЯ"),
            
            // Fallback for missing values
            ("{}", "{}")
        };
    }
}