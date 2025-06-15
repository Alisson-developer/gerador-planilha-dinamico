package com.snowplat.fabrica_planilha.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.context.support.PropertySourcesPlaceholderConfigurer;

@Configuration
@PropertySource("classpath:messages.properties")
public class PlaceholderConfig {

    /**
     * Bean obrigatório para permitir que o Spring processe placeholders (${...})
     * em valores de atributos de anotação (como @Operation(description = "...")).
     * Ele carrega automaticamente os arquivos .properties (application.properties
     * e messages.properties) para resolver ${chave} em tempo de inicialização.
     */
    @Bean
    public static PropertySourcesPlaceholderConfigurer propertySourcesPlaceholderConfigurer() {
        return new PropertySourcesPlaceholderConfigurer();
    }

}
