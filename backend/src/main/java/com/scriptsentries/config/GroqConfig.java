package com.scriptsentries.config;

import org.springframework.ai.openai.OpenAiChatModel;
import org.springframework.ai.openai.OpenAiChatOptions;
import org.springframework.ai.openai.api.OpenAiApi;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class GroqConfig {
    @Bean
    public OpenAiChatModel openAiChatModel(@Value("${spring.ai.openai.api-key}") String apiKey) {
        // We hardcode the URL here to bypass all discovery bugs
        var openAiApi = new OpenAiApi("https://api.groq.com/openai", apiKey);
        return new OpenAiChatModel(openAiApi, OpenAiChatOptions.builder()
                .withModel("llama-3.3-70b-versatile")
                .withTemperature(0.1f)
                .build());
    }
}