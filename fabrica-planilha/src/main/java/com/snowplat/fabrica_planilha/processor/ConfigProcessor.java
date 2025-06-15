package com.snowplat.fabrica_planilha.processor;

import com.fasterxml.jackson.databind.JsonNode;
import java.lang.reflect.Method;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Classe responsável por processar dinamicamente a geração de planilhas XLSX com base em um {@link com.fasterxml.jackson.databind.JsonNode}.
 * <p>
 * Seu funcionamento é baseado em reflexão (<code>Java Reflection</code>), permitindo que novas configurações
 * sejam facilmente adicionadas sem a necessidade de modificar a lógica principal de despacho (dispatch).
 * </p>
 *
 * <p><strong>Responsabilidades principais:</strong></p>
 * <ol>
 *   <li>Receber um <code>JsonNode</code> contendo a configuração da planilha.</li>
 *   <li>Iterar dinamicamente sobre as chaves do JSON (ex.: <code>"workbook"</code>, <code>"aba"</code>, <code>"linha"</code>).</li>
 *   <li>Para cada chave, utilizar reflexão para invocar o método correspondente: <code>process{Chave}(JsonNode)</code>.</li>
 * </ol>
 *
 * <p><strong>Exemplo:</strong> se o JSON contiver a chave <code>"workbook"</code>, o método <code>workbook(JsonNode)</code> será chamado automaticamente.
 * Da mesma forma, para <code>"aba"</code>, será chamado <code>processAba(JsonNode)</code>, e assim por diante.</p>
 *
 * <p>Essa abordagem segue o princípio da extensão com mínima modificação (Open/Closed Principle),
 * permitindo que novos comportamentos sejam adicionados criando apenas novos métodos <code>processX(JsonNode)</code>.</p>
 */
public class ConfigProcessor {

    private final XSSFWorkbook workbook;
    private Sheet currentSheet;
    private int currentRowNum;

    public ConfigProcessor(XSSFWorkbook workbook) {
        this.workbook = workbook;
        this.currentSheet = null;
        this.currentRowNum = 0;
    }

    /**
     * Ponto de entrada principal para o processamento da configuração.
     * <p>
     * Percorre dinamicamente todas as chaves de primeiro nível do JSON raiz e,
     * para cada uma, tenta invocar o método correspondente <code>process{Chave}(JsonNode)</code> utilizando Java Reflection.
     * </p>
     *
     * <p>
     * Essa abordagem permite o desacoplamento entre o formato do JSON e a lógica de construção do arquivo XLSX,
     * facilitando a extensão da funcionalidade apenas com a adição de novos métodos <code>processX(JsonNode)</code>.
     * </p>
     *
     * @param rootNode Nó raiz do JSON contendo a configuração geral da planilha.
     */
    public void processar(JsonNode rootNode) {
        JsonValidator.validarCamposObrigatorios(rootNode);

        Iterator<String> fieldNames = rootNode.fieldNames();

        while (fieldNames.hasNext()) {
            String fieldName = fieldNames.next();
            JsonNode childNode = rootNode.get(fieldName);
            String methodName = "process" + capitalize(fieldName);

            try {
                Method method = this.getClass().getDeclaredMethod(methodName, JsonNode.class);
                method.setAccessible(true);
                method.invoke(this, childNode);

            } catch (NoSuchMethodException e) {
                System.err.printf("Chave \"%s\" não reconhecida. " +
                                "Implemente process%s(JsonNode) se necessário.%n",
                        fieldName, capitalize(fieldName));
            } catch (Exception e) {
                throw new RuntimeException(
                        String.format("Falha ao processar '%s': %s", fieldName, e.getMessage()), e);
            }

        }
    }

    /**
     * Processa a configuração de uma <strong>workbook</strong>, que representa um <code>Workbook</code> (arquivo XLSX).
     * <p>
     * O JSON esperado deve seguir o seguinte formato:
     * </p>
     *
     * <pre>{@code
     * {
     *   "workbook": {
     *     "nomeWorkbook": "MeuRelatorio",
     *     "abas": [ ... ] // array de configurações de abas (sheets)
     *   }
     * }
     * }</pre>
     *
     * <p>
     * A propriedade <code>nomeWorkbook</code> será usada como nome do arquivo <code>.xlsx</code> gerado.
     * A propriedade <code>abas</code> contém um array de objetos com a configuração de cada aba da planilha.
     * </p>
     *
     * @param workbookNode JsonNode que representa a estrutura da workbook, incluindo o nome e as abas do Workbook.
     */
    private void processWorkbook(JsonNode workbookNode) {
        JsonNode nomeNode = workbookNode.get("nomeWorkbook");
        JsonValidator.validarNoComoTexto(nomeNode, "nomeWorkbook");

        JsonValidator.validarNoComoArray(workbookNode, "abas");
        if (workbookNode.has("abas") && workbookNode.get("abas").isArray()) {
            for (JsonNode abaNode : workbookNode.get("abas")) {
                processAba(abaNode);
            }
        }
    }

    /**
     * Processa a configuração de uma <strong>aba</strong> (Sheet) da planilha.
     * <p>
     * O JSON esperado para cada aba deve seguir o formato:
     * </p>
     *
     * <pre>{@code
     * {
     *   "nomeAba": "Exemplo",
     *   "linha": {
     *     "isCabecalho": true,
     *     "campos": [
     *       "id",
     *       "nome",
     *       "valor"
     *     ]
     *   },
     *   "dados": [
     *     {
     *       "id": 1,
     *       "nome": "João",
     *       "valor": 100
     *     },
     *     {
     *       "id": 2,
     *       "nome": "Maria",
     *       "valor": 200
     *     }
     *   ]
     * }
     * }</pre>
     *
     * <p>
     * A propriedade <code>linha.isCabecalho</code> indica se a primeira linha deve ser tratada como cabeçalho.
     * O array <code>campos</code> define a ordem das colunas e os atributos a serem lidos dos objetos do array <code>dados</code>.
     * </p>
     *
     * @param abaNode JsonNode que contém as informações da aba: nome, campos e dados.
     */
    private void processAba(JsonNode abaNode) {
        JsonNode nomeAbaNode = abaNode.get("nomeAba");
        JsonValidator.validarNoComoTexto(nomeAbaNode, "nomeAba");
        String nomeAba = nomeAbaNode.asText();

        currentSheet = workbook.createSheet(nomeAba);
        currentRowNum = 0;

        if (abaNode.has("cabecalho")) {
            JsonNode cabecalhoNode = abaNode.get("cabecalho");
            processCabecalho(cabecalhoNode);
        } else {
            throw new IllegalArgumentException(
                    String.format("Uma aba (\"%s\") deve conter o nó \"cabecalho\".", nomeAba));
        }

        if (abaNode.has("dados")) {
            JsonValidator.validarNoComoArray(abaNode, "dados");
            for (JsonNode dadoNode : abaNode.get("dados")) {
                processDados(dadoNode);
            }
        }
    }

    /**
     * Processa a configuração de <strong>cabeçalho</strong> da planilha.
     * <p>
     * Atualmente, é considerado apenas o processamento da <em>linha de cabeçalho</em>.
     * O JSON esperado deve estar no seguinte formato:
     * </p>
     *
     * <pre>{@code
     * {
     *   "campos": [
     *     "id",
     *     "nome",
     *     "valor"
     *   ]
     * }
     * }</pre>
     *
     * <p>
     * O array <code>campos</code> define os nomes das colunas que serão exibidas.
     * </p>
     *
     * @param cabecalhoNode JsonNode que contém a configuração do cabeçalho e sues campos.
     */
    private void processCabecalho(JsonNode cabecalhoNode) {
        JsonValidator.validarNoComoArray(cabecalhoNode, "campos");
        JsonNode camposNode = cabecalhoNode.get("campos");

        CellStyle estiloAplicado = null;
        if (cabecalhoNode.has("estilo")) {
            JsonNode estiloNode = cabecalhoNode.get("estilo");
            JsonValidator.validarNoObjetoDireto(estiloNode, "cabecalho.estilo");

            if (estiloNode.has("alinhamento")) {
                JsonNode alin = estiloNode.get("alinhamento");
                JsonValidator.validarNoObjetoDireto(alin, "cabecalho.estilo.alinhamento");
                if (alin.has("horizontal")) {
                    JsonValidator.validarValorAlinhamento(
                            alin.get("horizontal").asText(), "horizontal");
                }
                if (alin.has("vertical")) {
                    JsonValidator.validarValorAlinhamento(
                            alin.get("vertical").asText(), "vertical");
                }
            }

            if (estiloNode.has("corTexto")) {
                JsonValidator.validarCorRGB(estiloNode.get("corTexto"), "cabecalho.estilo.corTexto");
            }
            if (estiloNode.has("corFundo")) {
                JsonValidator.validarCorRGB(estiloNode.get("corFundo"), "cabecalho.estilo.corFundo");
            }
            if (estiloNode.has("negrito")) {
                JsonValidator.validarNoComoBoolean(estiloNode, "cabecalho.estilo.negrito");
            }
            if (estiloNode.has("italico")) {
                JsonValidator.validarNoComoBoolean(estiloNode, "cabecalho.estilo.italico");
            }
            if (estiloNode.has("sublinhado")) {
                JsonValidator.validarNoComoBoolean(estiloNode, "cabecalho.estilo.sublinhado");
            }

            JsonValidator.validarBordas(estiloNode);

            estiloAplicado = aplicarEstilo(estiloNode, workbook);
        }

        Row headerRow = currentSheet.createRow(currentRowNum++);
        CellStyle styleParaUsar = estiloAplicado;
        if (styleParaUsar == null) {
            styleParaUsar = workbook.createCellStyle();
            Font fontPadrao = workbook.createFont();
            fontPadrao.setBold(true);
            styleParaUsar.setFont(fontPadrao);
        }

        int colIndex = 0;
        for (JsonNode campo : camposNode) {
            JsonValidator.validarNoComoTexto(campo, "cabecalho.campos[" + colIndex + "]");
            Cell cell = headerRow.createCell(colIndex++);
            cell.setCellValue(campo.asText());
            cell.setCellStyle(styleParaUsar);
        }
    }

    /**
     * Processa um objeto de <strong>dados</strong> genérico, onde cada JSON representa uma linha da planilha.
     * <p>
     * Para cada campo presente no JSON, o valor é lido e escrito sequencialmente nas células da linha atual da aba.
     * </p>
     *
     * <p>Exemplo de dado JSON esperado:</p>
     *
     * <pre>{@code
     * {
     *   "id": 1,
     *   "nome": "João",
     *   "valor": 100
     * }
     * }</pre>
     *
     * <p>
     * As chaves do objeto representam os nomes das colunas e os valores representam o conteúdo das células.
     * A ordem de escrita segue a ordem definida previamente em <code>linha.campos</code>.
     * </p>
     *
     * @param dadoNode JsonNode representando um objeto com atributos no formato chave/valor.
     */
    private void processDados(JsonNode dadoNode) {
        if (currentSheet == null) {
            throw new IllegalStateException(
                    "Tentativa de adicionar dados sem criar aba antes (processAba não foi invocado).");
        }
        JsonValidator.validarNoObjetoDireto(dadoNode, "dado");

        final String ESTILO = "estilo";
        Row dataRow = currentSheet.createRow(currentRowNum++);
        int cellIndex = 0;

        Iterator<String> fieldNames = dadoNode.fieldNames();
        while (fieldNames.hasNext()) {
            String nomeCampo = fieldNames.next();

            if (nomeCampo.equals(ESTILO)) {
                continue;
            }

            JsonNode valorNo = dadoNode.get(nomeCampo);

            if (valorNo == null || valorNo.isNull()) {
                throw new IllegalArgumentException(
                        String.format("Campo \"%s\" em dados está ausente ou nulo.", nomeCampo));
            }

            Cell cell = dataRow.createCell(cellIndex++);

            if (dadoNode.has("estilo")) {
                JsonNode estiloNode = dadoNode.get("estilo");
                JsonValidator.validarNoObjetoDireto(estiloNode, "dado.estilo");

                if (estiloNode.has("alinhamento")) {
                    JsonNode alin = estiloNode.get("alinhamento");
                    JsonValidator.validarNoObjetoDireto(alin, "dado.estilo.alinhamento");
                    if (alin.has("horizontal")) {
                        JsonValidator.validarValorAlinhamento(
                                alin.get("horizontal").asText(), "horizontal");
                    }
                    if (alin.has("vertical")) {
                        JsonValidator.validarValorAlinhamento(
                                alin.get("vertical").asText(), "vertical");
                    }
                }

                if (estiloNode.has("corTexto")) {
                    JsonValidator.validarCorRGB(estiloNode.get("corTexto"), "dado.estilo.corTexto");
                }
                if (estiloNode.has("corFundo")) {
                    JsonValidator.validarCorRGB(estiloNode.get("corFundo"), "dado.estilo.corFundo");
                }
                if (estiloNode.has("negrito")) {
                    JsonValidator.validarNoComoBoolean(estiloNode, "dado.estilo.negrito");
                }
                if (estiloNode.has("italico")) {
                    JsonValidator.validarNoComoBoolean(estiloNode, "dado.estilo.italico");
                }
                if (estiloNode.has("sublinhado")) {
                    JsonValidator.validarNoComoBoolean(estiloNode, "dado.estilo.sublinhado");
                }

                JsonValidator.validarBordas(estiloNode);

                CellStyle estiloParaCelula = aplicarEstilo(estiloNode, workbook);
                cell.setCellStyle(estiloParaCelula);
            }

            if (valorNo.isNumber()) {
                cell.setCellValue(valorNo.asDouble());
            } else if (valorNo.isBoolean()) {
                cell.setCellValue(valorNo.asBoolean());
            } else {
                cell.setCellValue(valorNo.asText());
            }
        }
    }

    /**
     * Método auxiliar que <strong>capitaliza</strong> a primeira letra de uma palavra.
     * <p>
     * Exemplo: <code>"workbook"</code> → <code>"Workbook"</code>
     * </p>
     *
     * @param texto Palavra a ser capitalizada.
     * @return Texto com a primeira letra em maiúscula e as demais inalteradas.
     */
    private String capitalize(String texto) {
        if (texto == null || texto.isEmpty()) {
            return texto;
        }
        return texto.substring(0, 1).toUpperCase() + texto.substring(1);
    }

    /**
     * Cria e retorna um {@link org.apache.poi.ss.usermodel.CellStyle} configurado de acordo com as propriedades
     * de estilo definidas em um {@link com.fasterxml.jackson.databind.JsonNode}.
     * <p>
     * O JSON de estilo pode conter as seguintes chaves (todas opcionais):
     * </p>
     * <ul>
     *   <li><strong>alinhamento</strong>: objeto com campos
     *     <ul>
     *       <li><code>horizontal</code> (String): "ESQUERDA", "CENTRO", "DIREITA" ou "GERAL" (padrão).</li>
     *       <li><code>vertical</code> (String): "CIMA", "CENTRO", "BAIXO" (padrão: "BAIXO").</li>
     *     </ul>
     *   </li>
     *   <li><strong>corTexto</strong>: objeto com campos <code>R</code>, <code>G</code>, <code>B</code> (inteiros 0–255).</li>
     *   <li><strong>corFundo</strong>: objeto com campos <code>R</code>, <code>G</code>, <code>B</code> (inteiros 0–255).</li>
     *   <li><strong>negrito</strong>: boolean (true = texto em negrito).</li>
     *   <li><strong>italico</strong>: boolean (true = texto em itálico).</li>
     *   <li><strong>sublinhado</strong>: boolean (true = texto sublinhado).</li>
     *   <li><strong>bordaEsquerda</strong>: boolean (true = aplica borda fina à esquerda).</li>
     *   <li><strong>bordaDireita</strong>: boolean (true = aplica borda fina à direita).</li>
     *   <li><strong>bordaCima</strong>: boolean (true = aplica borda fina no topo).</li>
     *   <li><strong>bordaBaixo</strong>: boolean (true = aplica borda fina na parte inferior).</li>
     *   <li><strong>bordaVertical</strong>: boolean (true = aplica bordas finas nas laterais esquerda e direita).</li>
     *   <li><strong>bordaHorizontal</strong>: boolean (true = aplica bordas finas no topo e na parte inferior).</li>
     *   <li><strong>bordaTotal</strong>: boolean (true = aplica borda fina em todos os quatro lados).</li>
     * </ul>
     *
     * <p>
     * A aplicação das bordas segue esta ordem de precedência:
     * <ol>
     *   <li>Se <code>bordaTotal</code> estiver presente e for <code>true</code>, todas as bordas recebem estilo {@link org.apache.poi.ss.usermodel.BorderStyle#THIN}.</li>
     *   <li>Senão, se <code>bordaVertical</code> for <code>true</code>, aplica borda fina apenas em <strong>esquerda</strong> e <strong>direita</strong>.</li>
     *   <li>Senão, se <code>bordaHorizontal</code> for <code>true</code>, aplica borda fina apenas em <strong>topo</strong> e <strong>inferior</strong>.</li>
     *   <li>Por fim, valores individuais de <code>bordaEsquerda</code>, <code>bordaDireita</code>, <code>bordaCima</code> e <code>bordaBaixo</code>
     *       serão aplicados conforme estejam configurados como <code>true</code>.</li>
     * </ol>
     * </p>
     *
     * @param estiloNode O {@link com.fasterxml.jackson.databind.JsonNode} que contém as propriedades de estilo.
     * @param workbook   A instância de {@link org.apache.poi.xssf.usermodel.XSSFWorkbook} usada para criar o estilo.
     * @return Um {@link org.apache.poi.ss.usermodel.CellStyle} configurado de acordo com o JSON de estilo.
     */
    private CellStyle aplicarEstilo(JsonNode estiloNode, XSSFWorkbook workbook) {
        CellStyle estilo = workbook.createCellStyle();
        XSSFFont fonte = workbook.createFont();

        if (estiloNode.has("alinhamento")) {
            JsonNode alinhamento = estiloNode.get("alinhamento");

            if (alinhamento.has("horizontal")) {
                String horiz = alinhamento.get("horizontal").asText().toUpperCase();
                estilo.setAlignment(switch (horiz) {
                    case "ESQUERDA" -> HorizontalAlignment.LEFT;
                    case "CENTRO"   -> HorizontalAlignment.CENTER;
                    case "DIREITA"  -> HorizontalAlignment.RIGHT;
                    default         -> HorizontalAlignment.GENERAL;
                });
            }

            if (alinhamento.has("vertical")) {
                String vert = alinhamento.get("vertical").asText().toUpperCase();
                estilo.setVerticalAlignment(switch (vert) {
                    case "CIMA"   -> VerticalAlignment.TOP;
                    case "CENTRO" -> VerticalAlignment.CENTER;
                    default       -> VerticalAlignment.BOTTOM;
                });
            }
        }

        if (estiloNode.has("corTexto")) {
            JsonNode cor = estiloNode.get("corTexto");
            int r = cor.path("R").asInt(0);
            int g = cor.path("G").asInt(0);
            int b = cor.path("B").asInt(0);
            XSSFColor xssfCorTexto = new XSSFColor(new java.awt.Color(r, g, b), new DefaultIndexedColorMap());
            fonte.setColor(xssfCorTexto);
        }

        if (estiloNode.has("corFundo")) {
            JsonNode cor = estiloNode.get("corFundo");
            int r = cor.path("R").asInt(0);
            int g = cor.path("G").asInt(0);
            int b = cor.path("B").asInt(0);
            XSSFColor xssfCorFundo = new XSSFColor(new java.awt.Color(r, g, b), new DefaultIndexedColorMap());
            estilo.setFillForegroundColor(xssfCorFundo);
            estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        if (estiloNode.has("negrito") && estiloNode.get("negrito").asBoolean()) {
            fonte.setBold(true);
        }

        if (estiloNode.has("italico") && estiloNode.get("italico").asBoolean()) {
            fonte.setItalic(true);
        }

        if (estiloNode.has("sublinhado") && estiloNode.get("sublinhado").asBoolean()) {
            fonte.setUnderline(FontUnderline.SINGLE.getByteValue());
        }

        boolean aplicaBordaEsquerda  = estiloNode.path("bordaEsquerda").asBoolean(false);
        boolean aplicaBordaDireita   = estiloNode.path("bordaDireita").asBoolean(false);
        boolean aplicaBordaCima      = estiloNode.path("bordaCima").asBoolean(false);
        boolean aplicaBordaBaixo     = estiloNode.path("bordaBaixo").asBoolean(false);
        boolean aplicaBordaVertical  = estiloNode.path("bordaVertical").asBoolean(false);
        boolean aplicaBordaHorizontal= estiloNode.path("bordaHorizontal").asBoolean(false);
        boolean aplicaBordaTotal     = estiloNode.path("bordaTotal").asBoolean(false);

        if (aplicaBordaTotal) {
            estilo.setBorderTop(BorderStyle.THIN);
            estilo.setBorderBottom(BorderStyle.THIN);
            estilo.setBorderLeft(BorderStyle.THIN);
            estilo.setBorderRight(BorderStyle.THIN);
        } else {
            if (aplicaBordaVertical) {
                estilo.setBorderLeft(BorderStyle.THIN);
                estilo.setBorderRight(BorderStyle.THIN);
            }
            if (aplicaBordaHorizontal) {
                estilo.setBorderTop(BorderStyle.THIN);
                estilo.setBorderBottom(BorderStyle.THIN);
            }
            if (aplicaBordaEsquerda) {
                estilo.setBorderLeft(BorderStyle.THIN);
            }
            if (aplicaBordaDireita) {
                estilo.setBorderRight(BorderStyle.THIN);
            }
            if (aplicaBordaCima) {
                estilo.setBorderTop(BorderStyle.THIN);
            }
            if (aplicaBordaBaixo) {
                estilo.setBorderBottom(BorderStyle.THIN);
            }
        }

        estilo.setFont(fonte);
        return estilo;
    }


}
