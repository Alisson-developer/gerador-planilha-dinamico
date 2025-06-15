import { useState, type SetStateAction } from 'react'
import { Prism as SyntaxHighlighter } from 'react-syntax-highlighter';
import exempleJson from './assets/request-body.json';
import requestExampleBody from './assets/request-body-example.json';
import { okaidia } from 'react-syntax-highlighter/dist/esm/styles/prism';
import { Controlled as CodeMirror } from 'react-codemirror2';
import 'codemirror/lib/codemirror.css';
import 'codemirror/theme/dracula.css';
import 'codemirror/mode/javascript/javascript';
import './App.css'

const exampleJson = JSON.stringify(exempleJson, null, 2);
const requestExample = JSON.stringify(requestExampleBody, null, 2);

export default function App() {

  const [inputJson, setInputJson] = useState<string>(requestExample);

  const generateSpreadsheet = async () => {
    try {
      const response = await fetch('http://localhost:8080/service/spreadsheet/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: inputJson
      });
      if (!response.ok) throw new Error('Erro ao gerar planilha');
      const blob = await response.blob();

      const disposition = response.headers.get('Content-Disposition');

      let fileName = 'planilha.xlsx';

      if (disposition) {
        let match = disposition.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
        if (match && match[1]) {
          fileName = decodeURIComponent(match[1]);
        } else {
          match = disposition.match(/filename="?([^\";]+)"?/i);
          if (match && match[1]) {
            fileName = match[1].replace(/^=\?UTF-8\?Q\?(.*)\?=$/, '$1');
          }
        }
      }

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      alert('Falha ao gerar planilha');
    }
  };

  return (
    <div>
      <div className="space-y-15px">
        <div className="cabecalho-verde">
          <h1>Gerador de Planilhas</h1>
        </div>

        <div>
          <div className="margin-bt-15">
            <div className="text-align-start">
              <span className="font-bold text-red-500">Atenção:</span>
            </div>
            <ul className="text-align-start">
              <li>Abaixo você pode colar o JSON que deseja converter em uma planilha Excel.</li>
              <li>Certifique-se de que o JSON esteja no formato correto, conforme o exemplo abaixo.</li>
              <li>Você pode usar o exemplo abaixo como referência.</li>
              <li>O JSON deve ser válido e seguir o formato esperado.</li>
              <li>A planilha será gerada com base nos dados fornecidos.</li>
              <li>Aguarde a geração da planilha, isso pode levar alguns segundos dependendo do tamanho do JSON.</li>
              <li>Após a geração, a planilha será baixada automaticamente.</li>
              <li>Certifique-se de que o navegador permita downloads automáticos.</li>
            </ul>
            
            <p className="text-align-start">
              <span className="font-bold">Os campos permitidos são:</span>&nbsp;
              <span className="font-bold txt-green">
                "workbook", "nomeWorkbook", "abas", "nomeAba", "cabecalho", "campos", "estilo", "alinhamento", "horizontal", "vertical", "corTexto", "R", "G", "B", "corFundo", "negrito", "italico", "sublinhado", "bordaTotal", "bordaVertical", "bordaHorizontal", "bordaEsquerda", "bordaDireita", "bordaCima", "bordaBaixo", "dados"
              </span>
            </p>
            <p className="text-align-start">
              <span className="font-bold txt-red">Campos obrigatórios:</span>&nbsp;
              <span className="font-bold txt-orange">
                "workbook", "nomeWorkbook", "abas" (com pelo menos um item), "nomeAba" na primeira aba
              </span>
            </p>
            <span className="font-bold">Seu JSON</span>
          </div>
          <CodeMirror
            className="code-mirror"
            value={inputJson}
            options={{
              mode: {
                name: 'javascript',
                json: true
              },
              theme: 'dracula',
              lineNumbers: true,
              lint: true,
              tabSize: 2,
              autoCloseBrackets: true
            }}
            onBeforeChange={(_editor: any, _data: any, value: SetStateAction<string>) => setInputJson(value)}
          />
        </div>

        <div className="text-center">
          <button
            className="margin-bt-15 btn-green"
            onClick={generateSpreadsheet}
          >
            Gerar Planilha
          </button>
        </div>

        <div>
          <div>
            <span className="font-bold">JSON de Exemplo</span>
          </div>
          <SyntaxHighlighter language="json" style={okaidia} showLineNumbers>
            {exampleJson}
          </SyntaxHighlighter>
        </div>
      </div>
    </div>
  );
}
