using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using WebApplication3.Controllers;
using System.Text.RegularExpressions;

namespace WebApplication3.views {
    public class GeneratePdf  {
        BaseFont fonteBase = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, false);
        string caminho = @"nacional-logistica.png";
        double pesoL = 0;
        const int tamanhoFonteBase = 9;
        const string formulaRegex = @"(\d{6}\s)(.{40})\s*([\d\,\.]{1,15})[\sA-Z]*([\d\,\.]{1,15})[\sA-Z]*([\d\,\.]{1,15}) ([A-Z]{2}).*";

        Document pdf;
        MemoryStream arquivo;
        PdfWriter writer;
        byte[] byteArray;
        string[] fracionados;
        string[] congelados;

        public GeneratePdf(string[] fracionados = null, string[] congelados = null) {
            float pxPorMm = 72 / 25.2F;
            pdf = new Document(PageSize.A4, 7 * pxPorMm, 7 * pxPorMm,
                20 * pxPorMm, 7 * pxPorMm);
            arquivo = new MemoryStream();
            writer = PdfWriter.GetInstance(pdf, arquivo);

            pdf.Open();
            this.fracionados = fracionados;
            this.congelados = congelados;
        }

        ~GeneratePdf() {
            pdf = null;
            arquivo = null;
            writer = null;
            byteArray = null;
        }

        public FileStreamResult GeraPdf(string placa, string motorista, string ajudante, string dataSub, string empresa, ModeloOne[] notasDisp) {
            
            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph($"Romaneio de Clientes {empresa}\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            //Adição tabela
            var tabela = new PdfPTable(7);
            float[] larguraColunas = { 1.3f, 2.5f, 2f, 1.3f, 1.3f, 1.1f, 0.7f };
            tabela.SetWidths(larguraColunas);
            tabela.DefaultCell.BorderWidth = 0;
            tabela.WidthPercentage = 100;

            CriarCelulaTexto(tabela, "Nota Fiscal", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Cliente", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Bairro", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Cidade", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Valor (R$)", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Peso (KG)", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Vol", PdfPCell.ALIGN_CENTER, true);

            double peso = 0;
            double valor = 0;
            int volume = 0;

            foreach (var i in notasDisp) {
                peso += i.peso;
                valor += i.valorNf;
                volume += i.volume;
                CriarCelulaTexto(tabela, i.documento.ToString(), PdfPCell.ALIGN_CENTER);
                CriarCelulaTexto(tabela, i.cliente, PdfPCell.ALIGN_CENTER);
                CriarCelulaTexto(tabela, i.bairro, PdfPCell.ALIGN_CENTER);
                CriarCelulaTexto(tabela, i.cidade, PdfPCell.ALIGN_CENTER);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(i.valorNf, 2)), PdfPCell.ALIGN_CENTER);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(i.peso, 2)), PdfPCell.ALIGN_CENTER);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n0}", i.volume), PdfPCell.ALIGN_CENTER);

            }
            pesoL = peso;
            CriarCelulaTexto(tabela, $"Total:", PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, $"{notasDisp.Length} notas", PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, "", PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, "", PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(valor, 2)), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(peso, 2)), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n0}", volume), PdfPCell.ALIGN_CENTER);
            pdf.Add(tabela);

            pdf.Close();

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf2(string placa, string motorista, string ajudante, string dataSub, ModeloTwo[] Tdados) {
            

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nPor Clientes Frisa\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            foreach (var dados in Tdados) {
                var fontelb = new iTextSharp.text.Font(fonteBase, 7.9f,
                    iTextSharp.text.Font.ITALIC, BaseColor.BLACK);
                var label = new Paragraph(dados.dados + "\n------------------------" +
                    "---------------------------------------------------------------" +
                    "------------------------------", fontelb);
                label.Alignment = Element.ALIGN_LEFT;
                pdf.Add(label);
            }
            pdf.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");

        }

        public FileStreamResult GeraPdf3(string placa, string motorista, string ajudante, string dataSub, ModeloTwo[] Tdados) {
            

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nNão Fracionados Frisa\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            DataTable limpo = new DataTable();
            limpo.Columns.Add("Codigo", typeof(int));
            limpo.Columns.Add("Produto", typeof(string));
            limpo.Columns.Add("PesoLiquido", typeof(float));
            limpo.Columns.Add("Quantidade", typeof(int));
            limpo.Columns.Add("PesoBruto", typeof(float));
            limpo.Columns.Add("Unidade", typeof(string));

            var tabela1 = new PdfPTable(5);
            float[] larguraColunas1 = { 0.5f, 2.7f, 1f, 1f, 1.3f };
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase+1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso L", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Vol", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso B", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            pdf.Add(tabela1);

            // CONTEUDO DO PDF \\
            var enc = new Regex(formulaRegex);
            var ordemEntregaRex = new Regex(@"Ordem Entrega:(.*)");
            
            float qntTotal = 0;
            float pesoTotal = 0;

            foreach (var dados in Tdados) {
                var OrdemEntrega = ordemEntregaRex.Match(dados.dados).Groups[1].Value.Trim();
                var dados1 = enc.Replace(dados.dados, "######").Split("######");
                var format = enc.Matches(dados.dados);

                for (var j = 0; j < format.Count; j++) {

                    var codigo = format[j].Groups[1].Value;
                    var produto = format[j].Groups[2].Value.ToString().Trim();
                    var pesoLiquido = FormataString(format[j].Groups[3].Value);
                    var quantidade = FormataString(format[j].Groups[4].Value);
                    var pesoBruto = FormataString(format[j].Groups[5].Value);
                    var unidade = format[j].Groups[6].Value;
                    var dados2 = dados1[j + 1].Replace("\n", "").Trim();

                    qntTotal += float.Parse(quantidade);
                    pesoTotal += float.Parse(pesoBruto);


                    if (float.Parse(quantidade) == 1f && VerificaFrac(codigo)) {
                        dados2 = pesoLiquido;
                    }

                    if (dados2 == "") {

                        DataRow[] rows = limpo.Select("Codigo=" + codigo);
                        if (rows.Length == 0) {
                            DataRow linha = limpo.NewRow();
                            linha["Codigo"] = int.Parse(codigo);
                            linha["Produto"] = produto;
                            linha["PesoLiquido"] = float.Parse(pesoLiquido);
                            linha["Quantidade"] = (int)float.Parse(quantidade);
                            linha["PesoBruto"] = float.Parse(pesoBruto);
                            linha["Unidade"] = unidade;
                            limpo.Rows.Add(linha);
                        }
                        else {
                            var PesoLiquidoOK = rows[0].ItemArray[2].ToString();
                            var QuantidadeOK = rows[0].ItemArray[3].ToString();
                            var PesoBrutoOK = rows[0].ItemArray[4].ToString();
                            rows[0].BeginEdit();
                            rows[0]["PesoLiquido"] = float.Parse(PesoLiquidoOK) + float.Parse(pesoLiquido);
                            rows[0]["Quantidade"] = float.Parse(QuantidadeOK) + float.Parse(quantidade);
                            rows[0]["PesoBruto"] = float.Parse(PesoBrutoOK) + float.Parse(pesoBruto);
                            rows[0].EndEdit();
                        }
                    }
                }
            }
            var quantidadeTotal = 0f;
            var pesoBrutoTotal = 0f;

            for (int i = 0; i < limpo.Rows.Count; i++) {
                var tabela = new PdfPTable(5);
                float[] larguraColunas = { 0.5f, 2.7f, 1f, 1f, 1.3f };
                tabela.SetWidths(larguraColunas);
                tabela.DefaultCell.BorderWidth = 0;
                tabela.WidthPercentage = 100;

                quantidadeTotal += int.Parse(limpo.Rows[i][3].ToString());
                pesoBrutoTotal += float.Parse(limpo.Rows[i][4].ToString());

                var codigo = limpo.Rows[i][0].ToString();
                var produto = limpo.Rows[i][1].ToString();
                var pesoLiquido = limpo.Rows[i][2].ToString();
                var quantidade = limpo.Rows[i][3].ToString();
                var pesoBruto = limpo.Rows[i][4].ToString();
                var unidade = limpo.Rows[i][5].ToString();

                CriarCelulaTexto(tabela, codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase+1, italico: true, negrito: true);
                CriarCelulaTexto(tabela, produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(pesoLiquido)) + " " + unidade, PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", int.Parse(quantidade)) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(pesoBruto)) + " " + unidade + " BRUTO", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                pdf.Add(tabela);
            }

            var tabelaF = new PdfPTable(5);
            float[] larguraColunasF = { 0.5f, 2.7f, 1f, 1f, 1.3f };
            tabelaF.SetWidths(larguraColunasF);
            tabelaF.DefaultCell.BorderWidth = 0;
            tabelaF.WidthPercentage = 100;
            CriarCelulaTexto(tabelaF, "TOTAL:", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", quantidadeTotal) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"{String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", pesoBrutoTotal)} KG BRUTO", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);

            pdf.Add(tabelaF);

            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf4(string placa, string motorista, string ajudante, string dataSub, ModeloTwo[] Tdados) {
           

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nFracionados Frisa\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela1 = new PdfPTable(6);
            float[] larguraColunas1 = { 0.5f, 0.6f, 2.7f, 1f, 1f, 1.3f };
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Ordem", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase+1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso L", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Vol", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso B", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            pdf.Add(tabela1);

            // CONTEUDO DO PDF \\
            var enc = new Regex(formulaRegex);
            var ordemEntregaRex = new Regex(@"Ordem Entrega:(.*)");
            
            float qntTotal = 0;
            float pesoTotal = 0;
            foreach (var dados in Tdados) {
                var OrdemEntrega = ordemEntregaRex.Match(dados.dados).Groups[1].Value.Trim();
                var dados1 = enc.Replace(dados.dados, "######").Split("######");
                var format = enc.Matches(dados.dados);

                for (var j = 0; j < format.Count; j++) {

                    var codigo = format[j].Groups[1].Value;
                    var produto = format[j].Groups[2].Value.ToString().Trim();
                    var pesoLiquido = format[j].Groups[3].Value.Replace(",",".");
                    var quantidade = format[j].Groups[4].Value.Replace(",", ".");
                    var pesoBruto = format[j].Groups[5].Value.Replace(",", ".");
                    var unidade = format[j].Groups[6].Value;
                    var dados2 = dados1[j + 1].Replace("\n", "").Trim();

                    if (float.Parse(quantidade) == 1f && VerificaFrac(codigo)) {
                        dados2 = pesoLiquido;
                    }

                    if (dados2 != "") {

                        qntTotal += float.Parse(quantidade);
                        pesoTotal += float.Parse(pesoBruto);

                        var tabela = new PdfPTable(6);
                        float[] larguraColunas = { 0.3f, 0.6f, 2.7f, 1f, 1f, 1.3f };
                        tabela.SetWidths(larguraColunas);
                        tabela.DefaultCell.BorderWidth = 0;
                        tabela.WidthPercentage = 100;

                        CriarCelulaTexto(tabela, OrdemEntrega, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                        CriarCelulaTexto(tabela, codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                        CriarCelulaTexto(tabela, produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                        CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(pesoLiquido)) + " " + unidade, PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                        CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(quantidade)) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                        CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(pesoBruto)) + " " + unidade + " BRUTO", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                        pdf.Add(tabela);
                    }

                    if (dados2.Split(" ").Length > 1) {
                        var space = new Regex(@"\s{1,1000}");
                        dados2 = space.Replace(dados2, " ");
                        var fontelb2 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase + 1,
                               iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                        var label2 = new Paragraph(dados2 + "\n\n", fontelb2);
                        label2.Alignment = Element.ALIGN_LEFT;
                        pdf.Add(label2);
                    }
                }
            }

            var tabelaF = new PdfPTable(5);
            float[] larguraColunasF = { 0.5f, 2.7f, 1f, 1f, 1.3f };
            tabelaF.SetWidths(larguraColunasF);
            tabelaF.DefaultCell.BorderWidth = 0;
            tabelaF.WidthPercentage = 100;
            CriarCelulaTexto(tabelaF, "TOTAL:", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", qntTotal) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"{String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", pesoTotal)} KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);

            pdf.Add(tabelaF);

            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf5(string placa, string motorista, string ajudante, string dataSub, ModeloTree[] Tprodutos) {
            

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nResfriados Saudali\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela1 = new PdfPTable(4);
            float[] larguraColunas1 = { 0.7f, 2.7f, 1.3f, 1.3f};
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Volume", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso Liquido", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            pdf.Add(tabela1);

            DataTable limpo = new DataTable();
            limpo.Columns.Add("Sequencia", typeof(int));
            limpo.Columns.Add("Codigo", typeof(string));
            limpo.Columns.Add("Produto", typeof(string));
            limpo.Columns.Add("Volume", typeof(int));
            limpo.Columns.Add("PesoLiquido", typeof(float));

            // CONTEUDO DO PDF \\

            int qntTotal = 0;
            float pesoTotal = 0;
            foreach (var produto in Tprodutos) {
                if (!VerificaFrac(produto.codigo.Split("-")[0]) && !VerificaCong(produto.codigo.Split("-")[0])) {
                    DataRow[] rows = limpo.Select("Codigo=" + produto.codigo.Split("-")[0]);
                    if (rows.Length == 0) {
                        DataRow linha = limpo.NewRow();
                        linha["Sequencia"] = int.Parse(produto.sequencia);
                        linha["Codigo"] = produto.codigo.Split("-")[0].Trim();
                        linha["Produto"] = produto.produto.Trim();
                        linha["Volume"] = produto.volume;
                        linha["PesoLiquido"] = produto.peso;
                        limpo.Rows.Add(linha);
                    }
                    else {
                        var PesoLiquidoOK = rows[0].ItemArray[4].ToString();//.Replace(".", ",");
                        var QuantidadeOK = rows[0].ItemArray[3].ToString();
                        rows[0].BeginEdit();
                        rows[0]["PesoLiquido"] = float.Parse(PesoLiquidoOK) + produto.peso;
                        rows[0]["Volume"] = float.Parse(QuantidadeOK) + produto.volume;
                        rows[0].EndEdit();
                    }

                    qntTotal += produto.volume;
                    pesoTotal += produto.peso;

                }
            }
            for (int i = 0; i < limpo.Rows.Count; i++) {
                var tabela = new PdfPTable(4);
                float[] larguraColunas = { 0.7f, 2.7f, 1.3f, 1.3f };
                tabela.SetWidths(larguraColunas);
                tabela.DefaultCell.BorderWidth = 0;
                tabela.WidthPercentage = 100;

                var sequencia = limpo.Rows[i][0].ToString();
                var codigo = limpo.Rows[i][1].ToString();
                var produto = limpo.Rows[i][2].ToString();
                var volume = limpo.Rows[i][3].ToString();
                var pesoLiquido = limpo.Rows[i][4].ToString();

                CriarCelulaTexto(tabela, codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                CriarCelulaTexto(tabela, produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(volume)) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(pesoLiquido)) + " KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                pdf.Add(tabela);
            }

            var tabelaF = new PdfPTable(4);
            float[] larguraColunasF = { 0.7f, 2.7f, 1.3f, 1.3f };
            tabelaF.SetWidths(larguraColunasF);
            tabelaF.DefaultCell.BorderWidth = 0;
            tabelaF.WidthPercentage = 100;
            CriarCelulaTexto(tabelaF, "TOTAL:", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", qntTotal) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"{String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", pesoTotal)} KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);

            pdf.Add(tabelaF);

            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf6(string placa, string motorista, string ajudante, string dataSub, ModeloTree[] Tprodutos) {


            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nCongelados Saudali\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela1 = new PdfPTable(4);
            float[] larguraColunas1 = { 0.7f, 2.7f, 1.3f, 1.3f };
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Volume", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso Liquido", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            pdf.Add(tabela1);

            DataTable limpo = new DataTable();
            limpo.Columns.Add("Sequencia", typeof(int));
            limpo.Columns.Add("Codigo", typeof(string));
            limpo.Columns.Add("Produto", typeof(string));
            limpo.Columns.Add("Volume", typeof(int));
            limpo.Columns.Add("PesoLiquido", typeof(float));

            // CONTEUDO DO PDF \\

            int qntTotal = 0;
            float pesoTotal = 0;
            foreach (var produto in Tprodutos) {
                if (VerificaCong(produto.codigo.Split("-")[0])) {
                    DataRow[] rows = limpo.Select("Codigo=" + produto.codigo.Split("-")[0]);
                    if (rows.Length == 0) {
                        DataRow linha = limpo.NewRow();
                        linha["Sequencia"] = int.Parse(produto.sequencia);
                        linha["Codigo"] = produto.codigo.Split("-")[0].Trim();
                        linha["Produto"] = produto.produto.Trim();
                        linha["Volume"] = produto.volume;
                        linha["PesoLiquido"] = produto.peso;
                        limpo.Rows.Add(linha);
                    }
                    else {
                        var PesoLiquidoOK = rows[0].ItemArray[4].ToString();//.Replace(".", ",");
                        var QuantidadeOK = rows[0].ItemArray[3].ToString();
                        rows[0].BeginEdit();
                        rows[0]["PesoLiquido"] = float.Parse(PesoLiquidoOK) + produto.peso;
                        rows[0]["Volume"] = float.Parse(QuantidadeOK) + produto.volume;
                        rows[0].EndEdit();
                    }

                    qntTotal += produto.volume;
                    pesoTotal += produto.peso;

                }
            }
            for (int i = 0; i < limpo.Rows.Count; i++) {
                var tabela = new PdfPTable(4);
                float[] larguraColunas = { 0.7f, 2.7f, 1.3f, 1.3f };
                tabela.SetWidths(larguraColunas);
                tabela.DefaultCell.BorderWidth = 0;
                tabela.WidthPercentage = 100;


                var sequencia = limpo.Rows[i][0].ToString();
                var codigo = limpo.Rows[i][1].ToString();
                var produto = limpo.Rows[i][2].ToString();
                var volume = limpo.Rows[i][3].ToString();
                var pesoLiquido = limpo.Rows[i][4].ToString();

                CriarCelulaTexto(tabela, codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                CriarCelulaTexto(tabela, produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(volume)) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", float.Parse(pesoLiquido)) + " KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                pdf.Add(tabela);
            }

            var tabelaF = new PdfPTable(4);
            float[] larguraColunasF = { 0.7f, 2.7f, 1.3f, 1.3f };
            tabelaF.SetWidths(larguraColunasF);
            tabelaF.DefaultCell.BorderWidth = 0;
            tabelaF.WidthPercentage = 100;
            CriarCelulaTexto(tabelaF, "TOTAL:", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", qntTotal) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"{String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", pesoTotal)} KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);

            pdf.Add(tabelaF);

            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf7(string placa, string motorista, string ajudante, string dataSub, ModeloTree[] Tprodutos) {

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nFracionados Saudali\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase+1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela1 = new PdfPTable(5);
            float[] larguraColunas1 = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Sequência", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase+1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Volume", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso Liquido", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            pdf.Add(tabela1);

            // CONTEUDO DO PDF \\

            int qntTotal = 0;
            float pesoTotal = 0;
            foreach (var produto in Tprodutos) {
                
                if (VerificaFrac(produto.codigo.Split("-")[0])) {
                    
                    var tabela = new PdfPTable(5);
                    float[] larguraColunas = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
                    tabela.SetWidths(larguraColunas);
                    tabela.DefaultCell.BorderWidth = 0;
                    tabela.WidthPercentage = 100;

                    CriarCelulaTexto(tabela, produto.sequencia, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                    CriarCelulaTexto(tabela, produto.codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                    CriarCelulaTexto(tabela, produto.produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                    CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", produto.volume) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                    CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", produto.peso) + " KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                    pdf.Add(tabela);
                    qntTotal += produto.volume;
                    pesoTotal += produto.peso;

                }
            }

            var tabelaF = new PdfPTable(5);
            float[] larguraColunasF = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
            tabelaF.SetWidths(larguraColunasF);
            tabelaF.DefaultCell.BorderWidth = 0;
            tabelaF.WidthPercentage = 100;
            CriarCelulaTexto(tabelaF, "TOTAL:", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", qntTotal) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"{String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", pesoTotal)} KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);

            pdf.Add(tabelaF);

            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf8(
            int documento, 
            string cliente, 
            string bairro, 
            string cidade, 
            double valor,
            double peso,
            int volume,
            string dataSub,
            string motorista,
            string ajudante,
            string placa,
            ModeloTree[] Tprodutos) {
            
            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 20,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Relação Cliente\n\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            pdf.Add(titulo);

            //Imagem
            if (File.Exists(caminho)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminho);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }
            
            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela = new PdfPTable(7);
            float[] larguraColunas = { 1.3f, 2.5f, 2f, 1f, 1.3f, 1.1f, 0.7f };
            tabela.SetWidths(larguraColunas);
            tabela.DefaultCell.BorderWidth = 0;
            tabela.WidthPercentage = 100;

            CriarCelulaTexto(tabela, "Nota Fiscal", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Cliente", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Bairro", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Cidade", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Valor (R$)", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Peso (KG)", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Vol", PdfPCell.ALIGN_CENTER, true);

            CriarCelulaTexto(tabela, documento.ToString(), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, cliente, PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, bairro, PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, cidade, PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(valor, 2)), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(peso, 2)), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n0}", volume), PdfPCell.ALIGN_CENTER);

            pdf.Add(tabela);

            pdf.Add(new Paragraph("\n\n\n"));

            var tabela1 = new PdfPTable(5);
            float[] larguraColunas1 = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Sequência", PdfPCell.ALIGN_LEFT, negrito: true);
            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT,  negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT,  negrito: true);
            CriarCelulaTexto(tabela1, "Vol", PdfPCell.ALIGN_RIGHT,  negrito: true);
            CriarCelulaTexto(tabela1, "PesoB", PdfPCell.ALIGN_RIGHT,  negrito: true);

            foreach (var produto in Tprodutos) {
                CriarCelulaTexto(tabela1, produto.sequencia, PdfPCell.ALIGN_LEFT, italico: true, negrito: true);
                CriarCelulaTexto(tabela1, produto.codigo, PdfPCell.ALIGN_LEFT, italico: true, negrito: true);
                CriarCelulaTexto(tabela1, produto.produto, PdfPCell.ALIGN_LEFT, italico: true);
                CriarCelulaTexto(tabela1, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", produto.volume) + " CX", PdfPCell.ALIGN_RIGHT, italico: true);
                CriarCelulaTexto(tabela1, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", produto.peso) + " KG BRUTO", PdfPCell.ALIGN_RIGHT, italico: true);
            }
            pdf.Add(tabela1);

            pdf.Close();
            arquivo.Close();

            byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf9(
            int documento,
            string cliente,
            string bairro,
            string cidade,
            double valor,
            double peso,
            int volume,
            string dataSub,
            string motorista,
            string ajudante,
            string placa,
            string dados) {
            

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 25,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Relação Cliente\n\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            pdf.Add(titulo);

            //Imagem
            if (File.Exists(caminho)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminho);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, 11,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela = new PdfPTable(7);
            float[] larguraColunas = { 1.3f, 2.5f, 2f, 1f, 1.3f, 1.1f, 0.7f };
            tabela.SetWidths(larguraColunas);
            tabela.DefaultCell.BorderWidth = 0;
            tabela.WidthPercentage = 100;

            CriarCelulaTexto(tabela, "Nota Fiscal", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Cliente", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Bairro", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Cidade", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Valor (R$)", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Peso (KG)", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Vol", PdfPCell.ALIGN_CENTER, true);

            CriarCelulaTexto(tabela, documento.ToString(), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, cliente, PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, bairro, PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, cidade, PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(valor, 2)), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n}", Math.Round(peso, 2)), PdfPCell.ALIGN_CENTER);
            CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n0}", volume), PdfPCell.ALIGN_CENTER);

            pdf.Add(tabela);

            pdf.Add(new Paragraph("\n\n\n"));


            var fontelb = new iTextSharp.text.Font(fonteBase, 7.5f,
                    iTextSharp.text.Font.ITALIC, BaseColor.BLACK);
            var label = new Paragraph(dados, fontelb);
            label.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label);

            pdf.Close();
            arquivo.Close();

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf10(string placa, string motorista, string ajudante, string dataSub, ModeloTree[] Tprodutos) {

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nFracionados Saudali\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase + 1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            var tabela1 = new PdfPTable(5);
            float[] larguraColunas1 = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
            tabela1.SetWidths(larguraColunas1);
            tabela1.DefaultCell.BorderWidth = 0;
            tabela1.WidthPercentage = 100;

            CriarCelulaTexto(tabela1, "Sequência", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Codigo", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Produto", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Volume", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            CriarCelulaTexto(tabela1, "Peso Liquido", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
            pdf.Add(tabela1);

            // CONTEUDO DO PDF \\

            int qntTotal = 0;
            float pesoTotal = 0;
            foreach (var produto in Tprodutos) {

                if (VerificaFrac(produto.codigo.Split("-")[0])) {

                    var tabela = new PdfPTable(5);
                    float[] larguraColunas = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
                    tabela.SetWidths(larguraColunas);
                    tabela.DefaultCell.BorderWidth = 0;
                    tabela.WidthPercentage = 100;

                    CriarCelulaTexto(tabela, produto.sequencia, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                    CriarCelulaTexto(tabela, produto.codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, negrito: true);
                    CriarCelulaTexto(tabela, produto.produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                    CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", produto.volume) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                    CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", produto.peso) + " KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
                    pdf.Add(tabela);
                    qntTotal += produto.volume;
                    pesoTotal += produto.peso;

                }
            }

            var tabelaF = new PdfPTable(5);
            float[] larguraColunasF = { 0.8f, 0.7f, 2.7f, 1.3f, 1.3f };
            tabelaF.SetWidths(larguraColunasF);
            tabelaF.DefaultCell.BorderWidth = 0;
            tabelaF.WidthPercentage = 100;
            CriarCelulaTexto(tabelaF, "TOTAL:", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, "", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", qntTotal) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);
            CriarCelulaTexto(tabelaF, $"{String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", pesoTotal)} KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true);

            pdf.Add(tabelaF);

            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }

        public FileStreamResult GeraPdf11(string placa, string motorista, string ajudante, string dataSub, ModeloFour[] Orders) {

            //Titulo

            var fonteParagrafo = new iTextSharp.text.Font(fonteBase, 18,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            var titulo = new Paragraph("Romaneio de Separação\nPor Clientes Saudali\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT;
            titulo.SpacingAfter = 5;
            pdf.Add(titulo);

            //Imagem
            var caminhoImagem = caminho;

            if (File.Exists(caminhoImagem)) {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 50;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);
                var margemEsquerda = pdf.PageSize.Width - pdf.RightMargin - larguraLogo;
                var margemTopo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerda, margemTopo);
                writer.DirectContent.AddImage(logo, false);
            }

            var fontelb1 = new iTextSharp.text.Font(fonteBase, tamanhoFonteBase + 1,
                iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            string mens = $"Placa    : {(placa == "" ? "______________________" : placa)}\n" +
                $"Motorista: {(motorista == "" ? "______________________" : motorista)}\n" +
                $"Ajudante : {(ajudante == "" ? "______________________" : ajudante)}\n" +
                $"Data     : {(dataSub == "" ? "______________________" : dataSub)}\n\n";
            var label1 = new Paragraph(mens, fontelb1);
            label1.Alignment = Element.ALIGN_LEFT;
            pdf.Add(label1);

            
            // CONTEUDO DO PDF \\

            foreach (var order in Orders) {
                //Faz uma tabela com nome do cliente, sequencia e pedido

                var tabela1 = new PdfPTable(3);
                float[] larguraColunas1 = { 2.5f, 0.5f, 0.8f};
                tabela1.SetWidths(larguraColunas1);
                tabela1.DefaultCell.BorderWidth = 0;
                tabela1.WidthPercentage = 100;

                CriarCelulaTexto(tabela1, "Cliente: "+ order.cliente, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1);
                CriarCelulaTexto(tabela1, "NF: "+order.order, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1);
                CriarCelulaTexto(tabela1, "Sequencia: "+order.sequence, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1);

                pdf.Add(tabela1);

                //FAz uma tabela com os produtos sem 

                foreach (var box in order.products) {
                    var tabela = new PdfPTable(5);
                    float[] larguraColunas = { 0.5f, 0.8f, 4f, 1.3f, 1.3f };
                    tabela.SetWidths(larguraColunas);
                    tabela.DefaultCell.BorderWidth = 0;
                    tabela.WidthPercentage = 100;

                    CriarCelulaTexto(tabela, "         ", PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, BorderWidthTop:0);
                    CriarCelulaTexto(tabela, box.codigo, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, BorderWidthTop: 0);
                    CriarCelulaTexto(tabela, box.produto, PdfPCell.ALIGN_LEFT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, BorderWidthTop: 0);
                    CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", box.volume) + " CX", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, BorderWidthTop: 0);
                    CriarCelulaTexto(tabela, String.Format(new System.Globalization.CultureInfo("pt-BR"), "{0:n3}", box.peso) + " KG", PdfPCell.ALIGN_RIGHT, tamanhoFonte: tamanhoFonteBase + 1, italico: true, BorderWidthTop: 0);
                    pdf.Add(tabela);
                }
                pdf.Add(new Paragraph("\n"));
            }


            pdf.Close();
            arquivo.Close();

            pdf = null;
            writer = null;

            byte[] byteArray = arquivo.ToArray();

            arquivo = new MemoryStream();
            arquivo.Write(byteArray, 0, byteArray.Length);
            arquivo.Position = 0;

            return new FileStreamResult(arquivo, "application/pdf");
        }


        public string FormataString(string entrada) {
            return entrada.Split(",")[0].Replace(".", "") + "." + entrada.Split(",").Last();
        }


        private void CriarCelulaTexto(PdfPTable tabela,
           string texto, int alinhamentoHorz = PdfPCell.ALIGN_LEFT,
           bool negrito = false, bool italico = false,
           int tamanhoFonte = tamanhoFonteBase, int alturaCelula = 0, int BorderWidthTop = 1) {
            int estilo = iTextSharp.text.Font.NORMAL;
            if (negrito && italico) {
                estilo = iTextSharp.text.Font.BOLDITALIC;
            }
            else if (negrito) {
                estilo = iTextSharp.text.Font.BOLD;
            }
            else if (italico) {
                estilo = iTextSharp.text.Font.ITALIC;
            }
            var fonteCelula = new iTextSharp.text.Font(fonteBase, tamanhoFonte, estilo, BaseColor.BLACK);
            var celula = new PdfPCell(new Phrase(texto, fonteCelula));
            celula.HorizontalAlignment = alinhamentoHorz;
            celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            celula.Border = 0;
            celula.BorderWidthTop = BorderWidthTop;
            celula.BorderColorTop = BaseColor.GRAY;
            celula.MinimumHeight = alturaCelula;
            tabela.AddCell(celula);
        }
    
        private bool VerificaFrac(string codigo) {
            foreach (var codigoFrac in fracionados) {
                if (int.Parse(codigo) == int.Parse(codigoFrac)) return true;
            }
            return false;
        }
        private bool VerificaCong(string codigo) {
            foreach (var codigoCong in congelados) {
                if (int.Parse(codigo) == int.Parse(codigoCong)) return true;
            }
            return false;
        }
    }
}