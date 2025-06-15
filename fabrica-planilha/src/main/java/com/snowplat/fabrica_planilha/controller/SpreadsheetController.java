package com.snowplat.fabrica_planilha.controller;

import com.fasterxml.jackson.databind.JsonNode;
import com.snowplat.fabrica_planilha.service.SpreadsheetService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.enums.ParameterIn;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.ExampleObject;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.tags.Tag;
import java.nio.charset.StandardCharsets;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

/**
 *
 * Controlador REST responsável por expor endpoints para geração dinâmica de arquivos Excel (<code>.xlsx</code>)
 * a partir de configurações fornecidas via JSON.
 *
 * <p>
 * Esta classe define um endpoint HTTP para receber a estrutura de configuração de uma planilha (como nome do arquivo,
 * abas, colunas e dados), gerar o arquivo em memória e retornar diretamente o conteúdo como um arquivo baixável.
 * </p>
 *
 * <p>
 * A resposta gerada possui o cabeçalho <code>Content-Disposition</code> configurado com o nome do arquivo definido
 * no campo <code>"nomePasta"</code> do JSON, permitindo ao cliente o download direto do arquivo com o nome correto.
 * </p>
 *
 * <h3>Exemplo de requisição:</h3>
 * <pre>{@code
 * POST /service/spreadsheets/generate
 * Content-Type: application/json
 *
 * {
 *   "pasta": {
 *     "nomePasta": "RelatorioFinanceiro",
 *     "abas": [
 *       {
 *         "nomeAba": "Janeiro",
 *         "linha": {
 *           "isCabecalho": true,
 *           "campos": ["id", "nome", "valor"]
 *         },
 *         "dados": [
 *           { "id": 1, "nome": "João", "valor": 100 },
 *           { "id": 2, "nome": "Maria", "valor": 200 }
 *         ]
 *       }
 *     ]
 *   }
 * }
 * }</pre>
 *
 * <p>
 * A anotação {@code @RestController} indica que esta classe é um componente Spring que lida com requisições RESTful.
 * A anotação {@code @RequestMapping("/service/spreadsheets")} define a rota base dos endpoints do controlador.
 * </p>
 *
 * @see SpreadsheetService
 * @see com.fasterxml.jackson.databind.JsonNode
 * @see org.springframework.http.ResponseEntity
 * @see org.springframework.core.io.ByteArrayResource
 */
@Tag(name = "Fábrica de Planilha", description = "Endpoints para geração de planilhas Excel")
@RestController
@RequestMapping("/service/spreadsheet")
public class SpreadsheetController {

    private final SpreadsheetService spreadsheetService;

    public SpreadsheetController(SpreadsheetService spreadsheetService) {
        this.spreadsheetService = spreadsheetService;
    }

    /**
     * Gera um arquivo Excel (<code>.xlsx</code>) com base nas configurações fornecidas em um JSON
     * e retorna o conteúdo diretamente na resposta HTTP.
     * <p>
     * O nome do arquivo é definido pelo campo <code>"nomePasta"</code> presente no JSON,
     * sendo utilizado no cabeçalho <code>Content-Disposition</code> da resposta.
     * </p>
     *
     * <p><strong>Exemplo de JSON de entrada:</strong></p>
     * <pre>{@code
     * {
     *   "pasta": {
     *     "nomePasta": "RelatorioVendas",
     *     "abas": [
     *       { ... }
     *     ]
     *   }
     * }
     * }</pre>
     *
     * @param config Objeto JSON contendo a configuração da planilha e metadados do arquivo.
     * @return {@code ResponseEntity<byte[]>} com o arquivo gerado e os cabeçalhos apropriados para download.
     */
    @Operation(
            summary = "${doc.api.summary}",
            description = "${doc.api.description}",
            requestBody = @io.swagger.v3.oas.annotations.parameters.RequestBody(
                    description = "${doc.api.requestbody.description}",
                    required = true,
                    content = @Content(
                            mediaType = MediaType.APPLICATION_JSON_VALUE,
                            examples = {
                                    @ExampleObject(
                                            name = "<strong>Todos os campos utilizados no exemplo são os únicos campos disponíveis para uso.</strong>",
                                            summary = "Exemplo de payload completo para geração da planilha",
                                            externalValue = "/openapi-examples/request-body.json"
                                    )
                            }
                    )
            ),
            responses = {
                    @ApiResponse(
                            responseCode = "200",
                            description = "Arquivo .xlsx gerado com sucesso",
                            content = @Content(
                                    mediaType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    ),
                    @ApiResponse(
                            responseCode = "400",
                            description = "JSON inválido ou faltando campos obrigatórios",
                            content = @Content
                    )
            }
    )
    @PostMapping(
            value = "/generate",
            consumes = MediaType.APPLICATION_JSON_VALUE,
            produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    public ResponseEntity<StreamingResponseBody> generateSpreadsheet(
            @Parameter(
                    description = "Configuração e dados da planilha em formato JSON",
                    required = true,
                    in = ParameterIn.DEFAULT
            )
            @RequestBody JsonNode config) {

        JsonNode sanitizedConfig = spreadsheetService.sanitizeFormulas(config.deepCopy());
        String fileName = spreadsheetService.normalizaESanitizaNomeArquivo(config);

        byte[] excelBytes = spreadsheetService.generateWorkbookFromConfig(sanitizedConfig);

        StreamingResponseBody stream = os -> os.write(excelBytes);

        ContentDisposition contentDisposition = ContentDisposition
                .attachment()
                .filename(fileName, StandardCharsets.UTF_8)
                .build();

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition.toString())
                .contentType(MediaType
                        .parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .contentLength(excelBytes.length)
                .body(stream);
    }

}
