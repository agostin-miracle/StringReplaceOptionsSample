# StringReplaceOptionsSample

# 📝 Projeto: Automação de Documentos Word em C#

## 📖 Descrição

Este projeto, desenvolvido em **C#**, tem como objetivo automatizar a criação e edição de documentos **Word** (formato DOCX) de forma programática, sem a necessidade do Microsoft Word instalado.

Para isso, ele utiliza o componente **[Xceed Words for .NET](https://xceed.com/xceed-words-for-net/)**, uma biblioteca robusta que permite gerar, ler e manipular arquivos Word de maneira eficiente.

## 🛠️ Funcionalidades Principais

- Substituição dinâmica de textos e placeholders em documentos Word.
- Preenchimento automatizado de modelos com dados vindos de banco de dados ou APIs.
- Suporte à substituição em todo o documento, incluindo cabeçalhos, rodapés e tabelas.
- Customização do comportamento de substituição com **`StringReplaceTextOptions`**.

## 🔑 Uso da Classe `StringReplaceTextOptions`

A classe **`StringReplaceTextOptions`** do Xceed Words for .NET é utilizada para configurar de forma flexível a substituição de textos, possibilitando:

- Ignorar ou considerar maiúsculas/minúsculas.
- Substituir apenas a primeira ocorrência ou todas.
- Restringir a substituição a palavras inteiras ou trechos.
- Substituir múltiplos marcadores em diferentes partes do documento.

## 🚀 Benefícios

- **Automação**: elimina tarefas manuais de edição de documentos.
- **Escalabilidade**: gera relatórios, contratos e recibos em grande volume.
- **Flexibilidade**: personalização total do comportamento de substituição de textos.

## 📂 Exemplos de Aplicação

- Geração automática de contratos com dados personalizados.
- Emissão de recibos e certificados com informações dinâmicas.
- Criação de relatórios complexos com múltiplos placeholders.

## 💻 Exemplo de Uso em C# — Substituindo Vários Placeholders

```csharp
using System;
using System.Collections.Generic;
using Xceed.Words.NET;

namespace AutomacaoWord
{
    class Program
    {
        static void Main(string[] args)
        {
            // Caminho do arquivo modelo
            string caminhoArquivo = @"C:\Projetos\modelo.docx";

            // Dicionário de placeholders e seus valores
            var dados = new Dictionary<string, string>
            {
                { "{NOME_CLIENTE}", "Maria Oliveira" },
                { "{ENDERECO}", "Rua das Acácias, 123 - São Paulo/SP" },
                { "{DATA}", DateTime.Now.ToShortDateString() }
            };

            // Carrega o documento Word
            using (DocX documento = DocX.Load(caminhoArquivo))
            {
                // Percorre cada placeholder e substitui no documento
                foreach (var item in dados)
                {
                    var opcoes = new StringReplaceTextOptions
                    {
                        NewValue = item.Value,
                        ReplaceFirst = false, // substitui todas as ocorrências
                        MatchCase = false,    // ignora maiúsculas/minúsculas
                        WholeWord = false     // substitui mesmo dentro de outras palavras
                    };

                    documento.ReplaceText(item.Key, opcoes);
                }

                // Salva o documento modificado
                documento.SaveAs(@"C:\Projetos\modelo_preenchido.docx");
            }

            Console.WriteLine("Substituição múltipla concluída!");
        }
    }
}
