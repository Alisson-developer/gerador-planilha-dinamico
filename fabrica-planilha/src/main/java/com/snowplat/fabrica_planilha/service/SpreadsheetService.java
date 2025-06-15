package com.snowplat.fabrica_planilha.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.snowplat.fabrica_planilha.processor.ConfigProcessor;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.regex.Pattern;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

/**
 * Serviço responsável pela geração de arquivos Excel (<code>.xlsx</code>) com base em um JSON de configuração.
 * <p>
 * Esta classe encapsula a lógica de orquestração para criação de workbooks em memória, utilizando a biblioteca
 * <a href="https://poi.apache.org/">Apache POI</a> e o padrão de configuração baseado em {@link com.fasterxml.jackson.databind.JsonNode}.
 * </p>
 *
 * <p>
 * O principal método exposto por esta classe é {@link #generateWorkbookFromConfig(JsonNode)}, que recebe um objeto JSON com as
 * instruções para construir o arquivo Excel (como nome da pasta, abas, linhas e dados) e retorna o resultado como
 * um <code>byte[]</code> — ideal para resposta de endpoints REST ou persistência posterior.
 * </p>
 *
 * <p><strong>Exemplo de uso:</strong></p>
 * <pre>{@code
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
 * A lógica detalhada da interpretação do JSON é delegada à classe {@link ConfigProcessor}, que lida com a reflexão
 * e o mapeamento de chaves para métodos como <code>processPasta()</code>, <code>processAba()</code> etc.
 * </p>
 *
 * <p>
 * A anotação {@code @Service} indica que esta classe é gerenciada pelo Spring e pode ser injetada em controladores
 * ou outros serviços.
 * </p>
 *
 * @author Caos
 * @see ConfigProcessor
 * @see com.fasterxml.jackson.databind.JsonNode
 * @see org.apache.poi.xssf.usermodel.XSSFWorkbook
 */
@Service
public class SpreadsheetService {

    private static final Pattern FILENAME_SANITIZE = Pattern.compile("[^A-Za-z0-9_\\-]");

    /**
     * Gera um array de bytes (<code>byte[]</code>) representando um arquivo <code>.xlsx</code>,
     * com base nas configurações fornecidas em um {@link com.fasterxml.jackson.databind.JsonNode}.
     * <p>
     * Todo o processamento ocorre em memória: o arquivo <strong>não</strong> é gravado em disco.
     * Isso permite que o resultado seja retornado diretamente em um endpoint REST, por exemplo, via <code>ResponseEntity<byte[]></code>.
     * </p>
     *
     * @param config Objeto JSON contendo a configuração da planilha (workbook, abas, linhas e dados).
     * @return Arquivo <code>.xlsx</code> gerado em memória, como array de bytes.
     * @throws IOException Caso ocorra erro durante a escrita do arquivo na memória.
     */
    public byte[] generateWorkbookFromConfig(JsonNode config) {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            ConfigProcessor processor = new ConfigProcessor(workbook);
            processor.processar(config);

            workbook.write(outputStream);
            return outputStream.toByteArray();

        } catch (IOException e) {
            throw new RuntimeException("Erro ao gerar planilha Excel: " + e.getMessage(), e);
        }
    }

    /**
     * Recebe um array de bytes (<code>byte[]</code>) representando um arquivo <code>.xlsx</code>
     * e grava esse arquivo em disco no diretório <code>src/main/resources/templates/</code>.
     * <p>
     * Se o diretório <code>templates</code> não existir, ele será criado automaticamente.
     * O nome do arquivo será <code>nomeBase + ".xlsx"</code> (caso <code>nomeBase</code> já
     * contenha ".xlsx", não acrescenta outra extensão).
     * </p>
     *
     * @param excelBytes Array de bytes contendo todo o conteúdo do .xlsx gerado em memória.
     * @param nomeBase   Nome base para o arquivo (sem extensão ou com extensão ".xlsx").
     * @return Caminho absoluto do arquivo gerado em disco.
     * @throws RuntimeException Caso ocorra erro ao criar diretório ou ao escrever o arquivo.
     */
    public void saveWorkbookToLocal(byte[] excelBytes, String nomeBase) {
        String fileName = nomeBase.toLowerCase().endsWith(".xlsx")
                ? nomeBase
                : nomeBase + ".xlsx";

        Path templatesDir = Paths.get("src", "main", "resources", "templates");

        try {
            Files.createDirectories(templatesDir);
        } catch (IOException e) {
            throw new RuntimeException("Não foi possível criar o diretório 'resources/templates': " + e.getMessage(), e);
        }

        Path outputPath = templatesDir.resolve(fileName);

        try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
            fos.write(excelBytes);
            fos.flush();
        } catch (IOException e) {
            throw new RuntimeException("Erro ao salvar o arquivo Excel em '" +
                    outputPath + "': " + e.getMessage(), e);
        }
    }

    public String normalizaESanitizaNomeArquivo(JsonNode config) {
        JsonNode workbook = config.path("workbook");
        String rawName = workbook.path("nomeWorkbook").asText("WorkbookSemNome");
        String safeBase = sanitizeFileName(rawName);
        return safeBase.endsWith(".xlsx")
                ? safeBase
                : safeBase + ".xlsx";
    }

    /**
     * Remove caracteres perigosos do nome do arquivo para prevenir directory traversal
     * e outros ataques. Também corta comprimento para no máximo 64 caracteres.
     */
    private String sanitizeFileName(String input) {
        String cleaned = FILENAME_SANITIZE
                .matcher(input)
                .replaceAll("_");
        if (cleaned.length() > 64) {
            cleaned = cleaned.substring(0, 64);
        }
        return cleaned.isBlank() ? "Workbook" : cleaned;
    }

    /**
     * Percorre recursivamente o JSON procurando nós de fórmula e aplica whitelist simples:
     * - remove qualquer '=' no início
     * - permite apenas funções e operadores alfanuméricos, parênteses e operadores básicos.
     */
    public JsonNode sanitizeFormulas(JsonNode node) {
        if (node.isObject()) {
            ObjectNode obj = (ObjectNode) node;
            Iterator<String> fields = obj.fieldNames();
            while (fields.hasNext()) {
                String key = fields.next();
                JsonNode child = obj.get(key);
                if (child.isObject()
                        && "formula".equals(child.path("type").asText(null))
                        && child.has("expression")) {
                    String expr = child.get("expression").asText();

                    if (expr.startsWith("=")) {
                        expr = expr.substring(1);
                    }

                    expr = expr.replaceAll("[^A-Za-z0-9()+\\-*/^%,. ]", "");
                    ((ObjectNode) child).put("expression", expr);
                } else {
                    sanitizeFormulas(child);
                }
            }
        } else if (node.isArray()) {
            for (JsonNode item : node) {
                sanitizeFormulas(item);
            }
        }
        return node;
    }
}
