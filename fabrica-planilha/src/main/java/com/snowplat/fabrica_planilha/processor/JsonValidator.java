package com.snowplat.fabrica_planilha.processor;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.MissingNode;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Conjunto de validadores auxiliares para checar a estrutura do JsonNode
 * antes de processar cada parte (workbook, aba, cabeçalho, dados, estilo, etc.).
 */
public class JsonValidator {

    private static final Set<String> CAMPOS_VALIDOS = Set.of(
            "workbook", "nomeWorkbook", "abas", "nomeAba", "cabecalho", "campos", "estilo",
            "alinhamento", "horizontal", "vertical", "corTexto", "R", "G", "B", "corFundo",
            "negrito", "italico", "sublinhado",
            "bordaTotal", "bordaVertical", "bordaHorizontal", "bordaEsquerda", "bordaDireita", "bordaCima", "bordaBaixo",
            "dados"
    );

    /**
     * Valida se o JSON contém todos os campos obrigatórios para processar a estrutura da planilha.
     *
     * @param root O JsonNode raiz da requisição.
     * @throws IllegalArgumentException se algum campo obrigatório estiver ausente ou inválido.
     */
    public static void validarCamposObrigatorios(JsonNode root) {
        List<String> erros = new ArrayList<>();
        validarJson(root, "", erros);
        validarNoComoObjeto(root, "workbook");

        JsonNode workbookNode = root.get("workbook");
        validarCampoTexto(workbookNode, "nomeWorkbook");
        validarNoComoArray(workbookNode, "abas");

        JsonNode primeiraAba = workbookNode.get("abas").get(0);
        if (primeiraAba == null) {
            throw new IllegalArgumentException("A lista de 'abas' deve conter ao menos um item.");
        }

        validarCampoTexto(primeiraAba, "nomeAba");
    }


    /**
     * Valida recursivamente a estrutura de um nó JSON, verificando campos não permitidos
     * e detectando valores que parecem conter caminhos de arquivo. Erros encontrados são
     * acumulados na lista fornecida.
     * <p>
     * A validação percorre todos os elementos de objetos e arrays:
     * <ul>
     *   <li>Em objetos, itera sobre cada campo, constrói o caminho completo do campo
     *       e verifica:
     *     <ul>
     *       <li>Se o nome do campo está em {@code CAMPOS_VALIDOS}; caso contrário, registra erro.</li>
     *       <li>Se o valor textual contém padrões de caminho de arquivo (\\, / ou “..:”); caso afirmativo, registra erro.</li>
     *       <li>Chama recursivamente a si mesmo para validar subnós.</li>
     *     </ul>
     *   </li>
     *   <li>Em arrays, itera por índice e chama recursivamente a validação de cada elemento.</li>
     * </ul>
     *
     * @param node          o {@link JsonNode} atual a ser validado; não pode ser nulo
     * @param caminhoAtual  o caminho JSON construído até este nó (ex.: {@code "root.lista[2].campo"}).
     *                      Para o nó raiz, passar string vazia.
     * @param erros         lista mutável onde serão adicionadas descrições de todos os erros detectados;
     *                      deve estar inicializada antes da chamada e será preenchida com mensagens detalhadas.
     * @since 1.0
     */
    private static void validarJson(JsonNode node, String caminhoAtual, List<String> erros) {
        if (node.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> fields = node.fields();
            while (fields.hasNext()) {
                Map.Entry<String, JsonNode> field = fields.next();
                String nomeCampo = field.getKey();
                JsonNode valorCampo = field.getValue();
                String caminhoCompleto = caminhoAtual.isEmpty() ? nomeCampo : caminhoAtual + "." + nomeCampo;

                if (valorCampo.isTextual() && contemPath(valorCampo.asText())) {
                    erros.add("Valor com caminho de arquivo detectado em " + caminhoCompleto + ": " + valorCampo.asText());
                }

                if (nomeCampo.equals("cabecalho") || nomeCampo.equals("dados")) {
                    continue;
                }

                if (!CAMPOS_VALIDOS.contains(nomeCampo)) {
                    erros.add("Campo não permitido: " + caminhoCompleto);
                }

                validarJson(valorCampo, caminhoCompleto, erros);
            }
        } else if (node.isArray()) {
            for (int i = 0; i < node.size(); i++) {
                validarJson(node.get(i), caminhoAtual + "[" + i + "]", erros);
            }
        }

        if (!erros.isEmpty()) {
            throw new IllegalArgumentException("JSON inválido: " + erros);
        }
    }

    /**
     * Verifica se a string fornecida contém padrões que sugerem a presença de um caminho de arquivo,
     * tais como barra invertida ('\\'), barra normal ('/') ou sequência de duas pontuações e dois
     * pontos seguidos de dois pontos de unidade de volume de disco ("..:").
     *
     * @param valor  a string a ser examinada; não deve ser nula
     * @return {@code true} se o texto contiver qualquer ocorrência de '\\', '/' ou "..:";
     *         {@code false} caso contrário
     * @since 1.0
     */
    private static boolean contemPath(String valor) {
        return valor.contains("\\") || valor.contains("/") || valor.contains("..:");
    }

    /**
     * Assegura que o nó existe (não seja nulo ou MissingNode) e seja do tipo objeto.
     *
     * @param parentNode  O JsonNode que “deveria” conter o campo indicado.
     * @param fieldName   O nome do campo esperado dentro de parentNode.
     * @throws IllegalArgumentException Caso o campo não exista ou não seja um objeto JSON.
     */
    public static void validarNoComoObjeto(JsonNode parentNode, String fieldName) {
        if (parentNode == null || parentNode instanceof MissingNode || !parentNode.has(fieldName)) {
            throw new IllegalArgumentException(
                    String.format("Campo obrigatório \"%s\" ausente no JSON.", fieldName));
        }
        JsonNode filho = parentNode.get(fieldName);
        if (!filho.isObject()) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve ser um objeto (JsonNode.isObject()).", fieldName));
        }
    }

    /**
     * Valida se o campo existe e é um texto.
     *
     * @param parentNode O nó pai.
     * @param fieldName  O nome do campo.
     */
    public static void validarCampoTexto(JsonNode parentNode, String fieldName) {
        if (!parentNode.has(fieldName) || !parentNode.get(fieldName).isTextual()) {
            throw new IllegalArgumentException(
                    String.format("Campo obrigatório \"%s\" ausente ou não é uma string.", fieldName)
            );
        }
    }

    /**
     * Assegura que o nó existe (não seja nulo ou MissingNode) e seja do tipo array.
     *
     * @param parentNode  O JsonNode que “deveria” conter o campo indicado.
     * @param fieldName   O nome do campo esperado dentro de parentNode.
     * @throws IllegalArgumentException Caso o campo não exista ou não seja um array JSON.
     */
    public static void validarNoComoArray(JsonNode parentNode, String fieldName) {
        if (parentNode == null || parentNode instanceof MissingNode || !parentNode.has(fieldName)) {
            throw new IllegalArgumentException(
                    String.format("Campo obrigatório \"%s\" ausente no JSON ou é null.", fieldName));
        }
        JsonNode filho = parentNode.get(fieldName);
        if (!filho.isArray()) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve ser um array (JsonNode.isArray()).", fieldName));
        }
    }

    /**
     * Assegura que o JsonNode passado não seja nulo ou MissingNode, e retorne-o como objeto.
     *
     * @param node       O JsonNode a validar.
     * @param fieldName  Descrição do campo, para mensagens de erro.
     * @throws IllegalArgumentException Caso node seja null ou MissingNode ou não seja objeto.
     */
    public static void validarNoObjetoDireto(JsonNode node, String fieldName) {
        if (!node.isObject()) {
            throw new IllegalArgumentException(
                    String.format("Espere um objeto JSON para \"%s\".", fieldName));
        }
    }

    /**
     * Assegura que o JsonNode passado exista e seja do tipo texto (String).
     *
     * @param node       O JsonNode a validar.
     * @param fieldName  Descrição do campo, para mensagens de erro.
     * @throws IllegalArgumentException Caso node seja nulo, MissingNode ou não seja textual.
     */
    public static void validarNoComoTexto(JsonNode node, String fieldName) {
        if (!node.isTextual()) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve ser uma string (texto).", fieldName));
        }
    }

    /**
     * Assegura que o JsonNode passado exista e seja do tipo número.
     *
     * @param node       O JsonNode a validar.
     * @param fieldName  Descrição do campo, para mensagens de erro.
     * @throws IllegalArgumentException Caso node seja nulo, MissingNode ou não seja numérico.
     */
    public static void validarNoComoNumero(JsonNode node, String fieldName) {
        if (!node.isNumber()) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve ser numérico.", fieldName));
        }
    }

    /**
     * Assegura que o valor textual de alignment (horizontal ou vertical) esteja dentro dos valores permitidos.
     *
     * @param valor       O valor extraído do JSON (ex.: "CENTRO", "ESQUERDA", etc.).
     * @param tipo        "horizontal" ou "vertical", para mensagens de erro.
     * @throws IllegalArgumentException Caso valor não esteja entre os valores permitidos.
     */
    public static void validarValorAlinhamento(String valor, String tipo) {
        String v = valor.toUpperCase();
        if (tipo.equalsIgnoreCase("horizontal")) {
            if (!(v.equals("ESQUERDA") || v.equals("CENTRO") || v.equals("DIREITA") || v.equals("GERAL"))) {
                throw new IllegalArgumentException(
                        String.format("Alinhamento horizontal inválido: \"%s\". Use [ESQUERDA, CENTRO, DIREITA, GERAL].", valor));
            }
        } else if (tipo.equalsIgnoreCase("vertical")) {
            if (!(v.equals("CIMA") || v.equals("CENTRO") || v.equals("BAIXO"))) {
                throw new IllegalArgumentException(
                        String.format("Alinhamento vertical inválido: \"%s\". Use [CIMA, CENTRO, BAIXO].", valor));
            }
        } else {
            throw new IllegalArgumentException(
                    String.format("Tipo de alinhamento desconhecido: \"%s\". Deve ser 'horizontal' ou 'vertical'.", tipo));
        }
    }

    /**
     * Valida um objeto de cor RGB, assegurando que existam os campos R, G e B, que sejam números inteiros
     * entre 0 e 255.
     *
     * @param corNode    O JsonNode que representa o objeto de cor (deve ter R, G, B).
     * @param fieldLabel Para mensagens de erro (ex.: "corTexto", "corFundo").
     * @throws IllegalArgumentException Caso: corNode ausente, não objeto, ou algum canal RGB ausente ou fora de [0,255].
     */
    public static void validarCorRGB(JsonNode corNode, String fieldLabel) {
        String[] canais = { "R", "G", "B" };
        for (String canal : canais) {
            if (!corNode.has(canal)) {
                throw new IllegalArgumentException(
                        String.format("Canal \"%s\" ausente no objeto de cor \"%s\".", canal, fieldLabel));
            }
            JsonNode valorCanal = corNode.get(canal);
            if (!valorCanal.isInt()) {
                throw new IllegalArgumentException(
                        String.format("Canal \"%s\" em \"%s\" deve ser inteiro (0–255).", canal, fieldLabel));
            }
            int v = valorCanal.asInt();
            if (v < 0 || v > 255) {
                throw new IllegalArgumentException(
                        String.format("Valor do canal \"%s\" em \"%s\" fora do intervalo 0–255: %d.", canal, fieldLabel, v));
            }
        }
    }

    /**
     * Assegura que o JsonNode especificado seja um campo booleano válido.
     *
     * @param parentNode O nó pai que deveria ter o campo.
     * @param fieldName  O nome do campo booleano esperado.
     * @throws IllegalArgumentException Caso: campo ausente, não booleano.
     */
    public static void validarNoComoBoolean(JsonNode parentNode, String fieldName) {
        JsonNode filho = parentNode.get(fieldName);
        if (filho != null && !filho.isBoolean()) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve ser booleano (true/false).", fieldName));
        }
    }

    /**
     * Assegura que o JsonNode especificado contenha um campo numérico inteiro com nome fieldName.
     *
     * @param parentNode O nó pai que deveria ter o campo.
     * @param fieldName  O nome do campo inteiro esperado.
     * @throws IllegalArgumentException Caso: campo ausente, não inteiro.
     */
    public static void validarNoComoInteiro(JsonNode parentNode, String fieldName) {
        if (parentNode == null || parentNode instanceof MissingNode || !parentNode.has(fieldName)) {
            throw new IllegalArgumentException(
                    String.format("Campo inteiro \"%s\" ausente ou nulo.", fieldName));
        }
        JsonNode filho = parentNode.get(fieldName);
        if (!filho.isInt()) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve ser um inteiro.", fieldName));
        }
    }

    /**
     * Assegura que o JsonNode especificado contenha um campo textual que represente data (YYYY-MM-DD).
     *
     * Nota: esse método checa apenas se é texto e se bate o padrão básico, sem validar datas inválidas (ex.: 2025-13-40).
     *
     * @param parentNode O nó pai que deveria ter o campo.
     * @param fieldName  O nome do campo de data esperado.
     * @throws IllegalArgumentException Caso: campo ausente, não textual ou formato inválido (não bate \\d{4}-\\d{2}-\\d{2}).
     */
    public static void validarNoComoDataTexto(JsonNode parentNode, String fieldName) {
        JsonValidator.validarNoComoTexto(parentNode.get(fieldName), fieldName);
        String data = parentNode.get(fieldName).asText();
        if (!data.matches("\\d{4}-\\d{2}-\\d{2}")) {
            throw new IllegalArgumentException(
                    String.format("Campo \"%s\" deve estar no formato YYYY-MM-DD. Encontrado: %s", fieldName, data));
        }
    }

    /**
     * Assegura que o JsonNode especificado contenha campos de borda (booleanos), se presentes, e que sejam booleanos.
     *
     * @param estiloNode Nó que contém possíveis campos de borda.
     */
    public static void validarBordas(JsonNode estiloNode) {
        String[] bordas = {
                "bordaEsquerda", "bordaDireita", "bordaCima",
                "bordaBaixo", "bordaVertical", "bordaHorizontal", "bordaTotal"
        };
        for (String b : bordas) {
            if (estiloNode.has(b) && !estiloNode.get(b).isBoolean()) {
                throw new IllegalArgumentException(
                        String.format("Campo \"%s\" em estilo deve ser booleano (true/false).", b));
            }
        }
    }

}
