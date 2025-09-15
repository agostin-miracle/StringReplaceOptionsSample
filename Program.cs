using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection.Metadata;
using System.Threading.Channels;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ConsoleApp2
{

    internal class Program
    {
        static string STUFFED_STRING = "[_]";
        static void Main(string[] args)
        {
            /*
             * *
             *  Neste exemplo foi inserido uma sentença com duas palavras delimitados pela tag 'strong', esse delimitador STUFFED_STRING garante que a expressão será analisada dentro do requerimento da função
             *  posteriormente o mesmo delimitador STUFFED_STRING é removido para a impressão da sentença
             */
            string frase = "Nas <strong>culturas[_]africanas</strong>, o <i>zimbro</i> africano é considerado <b>sagrado</b> e usado em rituais religiosos e cerimônias de <em>purificação</em>.";


            var obj = new ChangeText();
            var changes = obj.CreateChanges(frase);

            //List<ChangeWords> changes = new List<ChangeWords>();
            //string[] palavras = frase.Split(' ');
            //foreach (string palavra in palavras)
            //{
            //    if (palavra.Contains("<b>") && palavra.Contains("</b>"))
            //        changes.Add(new ChangeWords { Text = ReplaceTag("b",palavra), Type = 1 });
            //    if (palavra.Contains("<i>") && palavra.Contains("</i>"))
            //        changes.Add(new ChangeWords { Text = ReplaceTag("i", palavra), Type = 2 });
            //    if (palavra.Contains("<strong>") && palavra.Contains("</strong>"))
            //        changes.Add(new ChangeWords { Text = ReplaceTag("strong", palavra), Type = 3 });
            //    if (palavra.Contains("<em>") && palavra.Contains("</em>"))
            //        changes.Add(new ChangeWords { Text = ReplaceTag("em", palavra), Type = 4 });

            //}

            /*
             * Lista as possiveis mudanças
             */
            foreach (ChangeWords item in changes)
            {
                Console.WriteLine($"palavra : {item.Text} tipo: {item.Type}");
            }

            /*
             * Executa as substituições
             */
            using (var document = DocX.Create("NewDocument.docx"))
            {
                var p = document.InsertParagraph();

                /*
                 * Remove todas as tags
                 */
                frase = obj.ReplaceTag(new string[] { "b", "strong", "i", "em" }, frase);

                /*
                 * Remove o delimitador, neste momento, a string ja foi apendada ao documento
                 */
                p.Append(frase.Replace(STUFFED_STRING, " "));

                foreach (ChangeWords item in changes)
                {
                    Formatting formatting = new Formatting();
                    formatting.Bold = true;

                    var options = new StringReplaceTextOptions
                    {
                        SearchValue = item.Text.Replace(STUFFED_STRING, " "),
                        NewValue = item.Text.Replace(STUFFED_STRING, " "),
                        NewFormatting = new Formatting()
                        {
                            Bold = (item.Type == 1 || item.Type == 3 || item.Type == 4) ? true : false,
                            Italic = item.Type == 2 ? true : false,
                            FontColor = item.Type == 4 ? Xceed.Drawing.Color.Blue : Xceed.Drawing.Color.Black
                        }

                    };
                    /*
                     * Executa a substituição
                     */
                    document.ReplaceText(options);

                }
                document.Save();
            }
            Console.WriteLine("Pressione uma tecla qualquer para continuar");
            Console.ReadKey();
        }


        //static string ReplaceTag(string[] tag, string text)
        //{
        //    if (tag.Length>0)
        //    {
        //        foreach (string s in tag)
        //        {
        //            text = ReplaceTag(s, text);
        //            //text = text.Replace(string.Format("<{0}>", s), "");
        //            //text = text.Replace(string.Format("</{0}>", s), "");
        //        }

        //    }
        //    return text;
        //}

        //static string ReplaceTag (string tag, string text)
        //{
        //    if (!string.IsNullOrEmpty(tag))
        //    {
        //        text = text.Replace(string.Format("<{0}>", tag), "");
        //        text = text.Replace(string.Format("</{0}>", tag), "");
        //    }
        //    return text;
        //}

    }

    /// <summary>
    /// Retêm os verbetes com marcação especial
    /// </summary>

    public class ChangeWords
    {
        /// <summary>
        /// Verbete
        /// </summary>
        public string Text { get; set; } = "";
        /// <summary>
        /// Tipo
        /// </summary>
        /// <list type="=bullet">
        ///<listheader>
        ///<term>Tipo</term>
        ///<description>Descrição</description>
        ///</listheader>
        ///<item><term>1</term><description>Bold</description></item>
        ///<item><term>2</term><description>Itálico</description></item>
        ///<item><term>3</term><description>Strong</description></item>
        ///<item><term>4</term><description>Enfático</description></item>
        ///</list>

        public byte Type { get; set; } = 0;
    }



    public class ChangeText
    {


        List<ChangeWords> Changes { get; set; } = new List<ChangeWords>();


        /// <summary>
        /// Efetua o parse de cada tag envolvida e anota para substituição
        /// </summary>
        /// <param name="frase">String contendo as marcações</param>
        /// <returns>List of ChangeWords</returns>
        public List<ChangeWords> CreateChanges(string frase)
        {
            string[] palavras = frase.Split(' ');
            foreach (string palavra in palavras)
            {
                if (palavra.Contains("<b>") && palavra.Contains("</b>"))
                    this.Changes.Add(new ChangeWords { Text = ReplaceTag("b", palavra), Type = 1 });
                if (palavra.Contains("<i>") && palavra.Contains("</i>"))
                    this.Changes.Add(new ChangeWords { Text = ReplaceTag("i", palavra), Type = 2 });
                if (palavra.Contains("<strong>") && palavra.Contains("</strong>"))
                    this.Changes.Add(new ChangeWords { Text = ReplaceTag("strong", palavra), Type = 3 });
                if (palavra.Contains("<em>") && palavra.Contains("</em>"))
                    this.Changes.Add(new ChangeWords { Text = ReplaceTag("em", palavra), Type = 4 });
            }
            return this.Changes;
        }


        public string ReplaceTag(string[] tag, string text)
        {
            if (tag.Length > 0)
            {
                foreach (string s in tag)
                {
                    text = ReplaceTag(s, text);
                    //text = text.Replace(string.Format("<{0}>", s), "");
                    //text = text.Replace(string.Format("</{0}>", s), "");
                }

            }
            return text;
        }

        public string ReplaceTag(string tag, string text)
        {
            if (!string.IsNullOrEmpty(tag))
            {
                text = text.Replace(string.Format("<{0}>", tag), "");
                text = text.Replace(string.Format("</{0}>", tag), "");
            }
            return text;
        }
    }
}

