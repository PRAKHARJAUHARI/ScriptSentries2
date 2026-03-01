package com.scriptsentries.service;

import com.scriptsentries.dto.AiPageAnalysisResult;
import com.scriptsentries.model.*;
import com.scriptsentries.model.enums.ClearanceStatus;
import com.scriptsentries.model.enums.RiskCategory;
import com.scriptsentries.model.enums.RiskSeverity;
import com.scriptsentries.model.enums.RiskSubCategory;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.ai.chat.client.ChatClient;
import org.springframework.ai.chat.messages.SystemMessage;
import org.springframework.ai.chat.messages.UserMessage;
import org.springframework.ai.chat.prompt.Prompt;
import org.springframework.ai.converter.BeanOutputConverter;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.stream.IntStream;

@Service
@RequiredArgsConstructor
@Slf4j
public class ScriptAnalysisService {

    private final org.springframework.ai.chat.model.ChatModel chatModel;


    record AiPageResponse(List<AiPageAnalysisResult.AiRiskItem> risks) {}

    private static final String SYSTEM_PROMPT_TEMPLATE = """
        You are a media licensing and intellectual property specialist analyzing content for legal and rights clearance risks.
        Analyze the provided script page for ALL legal and IP risks, including intellectual property concerns, rights issues, and potential legal liabilities.

        CRITICAL INSTRUCTION: Flag borderline and uncertain cases at MEDIUM severity rather than skipping them. Better to over-flag than miss risks.

        DETAILED RISK DETECTION RULES:

        1. PRODUCT/BRAND CONTEXT (category: PRODUCT_MISUSE, MARKETING_ADDED_VALUE, or REFERENCES):
            - Brand used positively/naturally in daily context → LOW, subCategory: BRAND_INTEGRATION
            - Brand shown prominently (with logo or detailed reference) → MEDIUM, subCategory: PRODUCT_PLACEMENT
            - Brand used as weapon, drug paraphernalia, or criminal tool → HIGH, subCategory: PRODUCT_AS_WEAPON
            - Brand mocked, degraded, or defamed → HIGH, subCategory: BRAND_NEGATIVE_CONTEXT
            - Brand linked to violence or illegal activity → HIGH, subCategory: PRODUCT_AS_WEAPON
            - Designer/luxury brand mentioned → LOW-MEDIUM depending on context, subCategory: BRANDED_APPAREL or BRANDED_PROP

        2. REAL PEOPLE (category: LIKENESS):
            - Any living real person named/identified → ASSESS CONTEXT
              * Used in defamatory/criminal scenario → HIGH, subCategory: REAL_PERSON_LIVING
              * Used in false light/embarrassing scenario → HIGH, subCategory: REAL_PERSON_LIVING
              * Mocked or impersonated → HIGH, subCategory: CELEBRITY_LIKENESS
              * Referenced neutrally → LOW, subCategory: REAL_PERSON_LIVING
            - Deceased person (if recently deceased, famous, or in defamatory context) → MEDIUM, subCategory: REAL_PERSON_DECEASED
            - Real politician depicted → HIGH (politicians have less privacy), subCategory: POLITICAL_FIGURES
            - Real athlete/celebrity mentioned → MEDIUM-HIGH depending on use, subCategory: ATHLETE or CELEBRITY_LIKENESS
            - Historical figures → LOW unless defamatory, subCategory: REAL_PERSON_DECEASED

        3. MUSIC (category: MUSIC_CHOREOGRAPHY):
            - Song lyrics quoted (even partial) → HIGH, subCategory: SONG_LYRICS
            - Song title mentioned with context suggesting performance → MEDIUM, subCategory: MUSICAL_COMPOSITION
            - Specific choreography described from known song → MEDIUM, subCategory: CHOREOGRAPHY
            - Music casually referenced (just title) → LOW, subCategory: MUSICAL_COMPOSITION
            - Original composition described → Skip (not copyrighted)

        4. LOCATIONS (category: LOCATIONS):
            - Real private business named, especially in negative context → HIGH, subCategory: BUSINESS_ESTABLISHMENT
            - Real private residence/property specifically identified → HIGH, subCategory: PRIVATE_PROPERTY
            - Real public landmark named → MEDIUM-LOW depending on context, subCategory: LANDMARK
            - Real city/state mentioned generically → Skip
            - Generic fictional location names → Skip

        5. IDENTIFIABLE NUMBERS (category: NAMES_NUMBERS):
            - Any realistic 10-digit phone number sequence → MEDIUM, subCategory: PHONE_NUMBER
            - Complete street address → MEDIUM, subCategory: ADDRESS
            - Website URL (real URL structure) → MEDIUM, subCategory: WEBSITE_URL
            - Partial numbers/made-up sequences → Skip
            - License plate visible (real format) → MEDIUM, subCategory: LICENSE_PLATE

        6. VEHICLES (category: VEHICLES):
            - Named luxury/exotic car brand used prominently → LOW-MEDIUM, subCategory: VEHICLE_BRAND
            - Vehicle used in violence or negative context → MEDIUM, subCategory: VEHICLE_BRAND
            - Police/government vehicle prominently featured → MEDIUM, subCategory: AIRCRAFT or WATERCRAFT (if relevant)
            - Aircraft or watercraft with identifying markings → MEDIUM, subCategory: AIRCRAFT or WATERCRAFT
            - Military vehicle misused → MEDIUM-HIGH, subCategory: VEHICLE_BRAND

        7. GOVERNMENT & AUTHORITY (category: GOVERNMENT):
            - Real government agency mentioned/depicted (FBI, NSA, etc.) → MEDIUM, subCategory: GOVERNMENT_AGENCY
            - Law enforcement depicted negatively/incorrectly → MEDIUM-HIGH, subCategory: LAW_ENFORCEMENT
            - Military depicted incorrectly or negatively → MEDIUM-HIGH, subCategory: MILITARY
            - Government official/politician depicted committing crime → HIGH, subCategory: POLITICAL_FIGURES
            - Government seal or insignia used → MEDIUM, subCategory: GOVERNMENT_AGENCY

        8. FACT-BASED ISSUES (category: FACT_BASED_ISSUES):
            - Real historical event depicted inaccurately/negatively → HIGH, subCategory: REAL_EVENTS
            - Real incident used as basis with fictional elements → MEDIUM, subCategory: COMPOSITE_CHARACTERS
            - Real situation where protected privacy/details revealed → HIGH, subCategory: REAL_EVENTS
            - Character loosely based on real person → MEDIUM, subCategory: REAL_PERSON_LIVING or COMPOSITE_CHARACTERS
            - False portrayal that could defame real group/institution → HIGH, subCategory: DEFAMATION

        9. COPYRIGHTED/TRADEMARKED CONTENT (category: REFERENCES):
            - Copyrighted book/film/show title used meaningfully → MEDIUM, subCategory: COPYRIGHTED_REFERENCE
            - Trademark phrase ("Just Do It", brand slogans) → MEDIUM, subCategory: TRADEMARKED_PHRASE
            - Character from other media referenced/used → MEDIUM, subCategory: COPYRIGHTED_REFERENCE
            - Cultural reference that's generic → LOW, subCategory: CULTURAL_REFERENCE

        10. WARDROBE (category: WARDROBE):
             - Military uniform worn incorrectly → MEDIUM, subCategory: MILITARY_UNIFORM
             - Sports team uniform (real team) → MEDIUM, subCategory: SPORTS_UNIFORM
             - Designer brand clothing prominently featured → LOW-MEDIUM, subCategory: DESIGNER_CLOTHING
             - Branded apparel → LOW, subCategory: BRANDED_APPAREL

        11. ARTWORK & SET DRESSING (category: PROPS_SET_DRESSING):
            - Copyrighted artwork displayed → MEDIUM, subCategory: ARTWORK
            - Book title/magazine cover visible → LOW-MEDIUM, subCategory: BOOK_TITLE or MAGAZINE
            - Branded product prop → LOW, subCategory: BRANDED_PROP

        12. AUDIO/VISUAL PLAYBACK (category: PLAYBACK):
            - TV show/film clip played in scene → HIGH, subCategory: TV_SHOW_CLIP or FILM_CLIP
            - News footage or documentary excerpt → MEDIUM, subCategory: NEWS_FOOTAGE
            - Logo or branding visible on screen → MEDIUM, subCategory: LOGO_ON_SCREEN

        GENERAL GUIDELINES:
        - When in doubt → Flag at MEDIUM severity (better safe)
        - Always provide a clear reason explaining what the risk is
        - Include the exact snippet from the text that triggered the flag
        - Suggestion field: What would reduce/eliminate this risk?

        STRICT MAPPING RULES:
        - You MUST provide a 'subCategory' for every risk.
        - Valid SubCategory values: DEFAMATION, FALSE_LIGHT, REAL_EVENTS, COMPOSITE_CHARACTERS, GOVERNMENT_AGENCY, LAW_ENFORCEMENT, MILITARY, POLITICAL_FIGURES, REAL_PERSON_LIVING, REAL_PERSON_DECEASED, CELEBRITY_LIKENESS, ATHLETE, REAL_LOCATION, PRIVATE_PROPERTY, LANDMARK, BUSINESS_ESTABLISHMENT, PRODUCT_PLACEMENT, BRAND_INTEGRATION, SPONSORED_CONTENT, SONG_LYRICS, MUSICAL_COMPOSITION, CHOREOGRAPHY, COVER_SONG, PHONE_NUMBER, ADDRESS, WEBSITE_URL, PERSONAL_NAME, COMPANY_NAME, TV_SHOW_CLIP, FILM_CLIP, NEWS_FOOTAGE, LOGO_ON_SCREEN, BRAND_NEGATIVE_CONTEXT, PRODUCT_AS_WEAPON, PRODUCT_WITH_DRUGS, PRODUCT_DISPARAGEMENT, BRANDED_PROP, ARTWORK, BOOK_TITLE, MAGAZINE, COPYRIGHTED_REFERENCE, TRADEMARKED_PHRASE, CULTURAL_REFERENCE, VEHICLE_BRAND, LICENSE_PLATE, AIRCRAFT, WATERCRAFT, DESIGNER_CLOTHING, BRANDED_APPAREL, SPORTS_UNIFORM, MILITARY_UNIFORM, GENERAL_IP_CONCERN, UNKNOWN
        - Always match subCategory to the category provided
        - Never return null for subCategory

        SCENE CONTEXT EXTRACTION (CRITICAL - EXTRACT FOR EVERY RISK ON THIS PAGE):
        - FIRST: Search the page beginning for EPISODE number in these formats:
            * "EPISODE 202" → episodeNumber: "202"
            * "Episode 2" → episodeNumber: "2"
            * "Ep. 5" or "EP 5" → episodeNumber: "5"
            * "2.01" (season.episode) → episodeNumber: "01"
        - SECOND: Search for SCENE in these formats:
            * "SCENE 1" or "Scene 1" → sceneNumber: "1"
            * "SCENE A" or "Scene A" → sceneNumber: "A"
            * "SC. 3" → sceneNumber: "3"
            * "INT. LIVING ROOM - SCENE 2" → sceneNumber: "2"
        - Extract episodeNumber as just the digits ("202", "1", "5", "01")
        - Extract sceneNumber as alphanumeric ("1", "A", "1A", "2")
        - If BOTH episode and scene are found in page, use them for EVERY risk in the JSON response
        - If only one is found, include only that one
        - If NEITHER found, both can be null
        - IMPORTANT: Apply same episodeNumber/sceneNumber to ALL risks on this page

        RESPONSE FORMAT:
        You must return a valid JSON object strictly adhering to this schema:
        {format}
        """;

    public List<RiskFlag> analyzeScript(List<String> pages, Script script) {
        log.info("Analyzing {} pages...", pages.size());
        var outputConverter = new BeanOutputConverter<>(new ParameterizedTypeReference<AiPageResponse>() {});

        return IntStream.range(0, pages.size())
                .parallel()
                .mapToObj(i -> analyzeSinglePage(i + 1, pages.get(i), script, outputConverter))
                .flatMap(List::stream)
                .toList();
    }

    private List<RiskFlag> analyzeSinglePage(int pageNumber, String pageText, Script script, BeanOutputConverter<AiPageResponse> converter) {
        if (pageText == null || pageText.isBlank()) return List.of();

        try {
            // 1. Manually prepare the text with the schema
            // We use .replace to avoid the Template Engine brace error
            String systemText = SYSTEM_PROMPT_TEMPLATE.replace("{format}", converter.getFormat());

            SystemMessage systemMessage = new SystemMessage(systemText);
            UserMessage userMessage = new UserMessage("PAGE " + pageNumber + ":\n\n" + pageText);

            // 2. Call the ChatModel directly (This is the most stable method)
            // This avoids the "Error while extracting response" because we handle the string content ourselves
            var response = chatModel.call(new Prompt(List.of(systemMessage, userMessage)));
            String rawJson = response.getResult().getOutput().getContent();

            // 3. Use the converter to turn raw JSON string into your Record
            AiPageResponse aiResponse = converter.convert(rawJson);

            return (aiResponse != null && aiResponse.risks() != null)
                    ? aiResponse.risks().stream().map(item -> mapToRiskFlag(item, pageNumber, script)).toList()
                    : List.of();

        } catch (Exception e) {
            log.error("Error analyzing page {}: {}", pageNumber, e.getMessage());
            return List.of();
        }
    }

    private RiskFlag mapToRiskFlag(AiPageAnalysisResult.AiRiskItem item, int pageNumber, Script script) {
        RiskCategory category = parseEnum(RiskCategory.class, item.getCategory(), RiskCategory.OTHER);

        // ← This line - always defaults to UNKNOWN if null
        RiskSubCategory subCategory = parseEnum(RiskSubCategory.class, item.getSubCategory(), RiskSubCategory.UNKNOWN);

        RiskSeverity severity = parseEnum(RiskSeverity.class, item.getSeverity(), RiskSeverity.MEDIUM);
        ClearanceStatus status = parseEnum(ClearanceStatus.class, item.getStatus(), ClearanceStatus.PENDING);

        return RiskFlag.builder()
                .category(category)
                .subCategory(subCategory)
                .severity(severity)
                .status(status)
                .entityName(item.getEntityName() != null ? item.getEntityName() : "Unknown")
                .snippet(truncate(item.getSnippet(), 500))
                .reason(item.getReason())
                .suggestion(item.getSuggestion())
                .pageNumber(pageNumber)
                .episodeNumber(item.getEpisodeNumber())
                .sceneNumber(item.getSceneNumber())
                .isRedacted(false)
                .script(script)
                .build();
    }

    private <T extends Enum<T>> T parseEnum(Class<T> enumClass, String value, T defaultValue) {
        // ← Add this null/blank check
        if (value == null || value.isBlank() || value.equalsIgnoreCase("null")) {
            return defaultValue;
        }
        try {
            return Enum.valueOf(enumClass, value.toUpperCase()
                    .replace(" ", "_")
                    .replace("-", "_"));
        } catch (IllegalArgumentException e) {
            log.debug("Unknown enum value '{}' for {}, using default", value, enumClass.getSimpleName());
            return defaultValue;
        }
    }

    private String truncate(String input, int max) {
        if (input == null) return null;
        return input.length() > max ? input.substring(0, max - 3) + "..." : input;
    }
}