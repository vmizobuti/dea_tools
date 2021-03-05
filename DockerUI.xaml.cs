using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;
using corel = Corel.Interop.VGCore;

namespace Dea_Tools
{
    public partial class DockerUI : UserControl
    {
        private corel.Application corelApp;
        private Styles.StylesController stylesController;
        public DockerUI(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);
            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }

        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }

        // Numerador Automático - Dea! Design
        private void numAutomatico(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Verifica se o usuário está selecionando objetos
            if (corelApp.ActiveSelectionRange.Count == 0)
            {
                System.Windows.MessageBox.Show("Selecione um objeto e/ou grupo para iniciar a numeração, já contendo um texto de base formatado.", "Atenção!");
            }
            else
            {
                // Verifica se os objetos selecionados incluem textos
                if (corelApp.ActiveSelection.Shapes.FindShapes(Type: corel.cdrShapeType.cdrTextShape).Count != 1)
                {
                    System.Windows.MessageBox.Show("Selecione um objeto/grupo que inclua um (apenas 1) texto para numeração.", "Atenção!");
                }
                else
                {
                    // Inicia o formulário para coleta das variáveis
                    FormNumerador formVal = new FormNumerador();

                    System.Windows.Forms.Application.Run(formVal);

                    // Inicia a ação do formulário após o usuário apertar 'Iniciar'
                    if (formVal.DialogResult == System.Windows.Forms.DialogResult.OK)
                    {
                        // Declara as variáveis de Início e Término da contagem
                        int numInicial = formVal.frmInicial;
                        int numFinal = formVal.frmFinal;

                        int rangeNum = numFinal - numInicial;

                        // Copia a seleção de objetos e altera a numeração e o posicionamento da cópia
                        for (int i = 0; i < rangeNum + 1; i++)
                        {
                            // Caso Apenas Pares tenha sido selecionado, pula a contagem ímpar
                            if (formVal.frmPares == true && ((numInicial + i) % 2) != 0)
                            {
                                continue;
                            }

                            // Caso Apenas Ímpares tenha sido selecionado, pula a contagem par
                            if (formVal.frmImpares == true && ((numInicial + i) % 2) == 0)
                            {
                                continue;
                            }

                            // Declara as variáveis de posicionamento da nova cópia
                            double offX1 = 0.0;
                            double offY1 = 0.0;
                            double offX2 = 0.0;
                            double offY2 = 0.0;

                            // Declara as variáveis de substituição do texto do grupo selecionado
                            string novoTexto = "";
                            string txtExistente = corelApp.ActiveSelection.Shapes.FindShape(Type: corel.cdrShapeType.cdrTextShape).Text.Story.Text;

                            // Caso tenha sido definido um prefixo, este é adicionado ao texto para substituição
                            if (formVal.frmPrefixo != "")
                            {
                                novoTexto += formVal.frmPrefixo;
                            }

                            // Adiciona o número '0' ao texto para substituição de acordo com o total de números da contagem
                            if (numFinal < 100 && formVal.frmZeros == true)
                            {
                                if (i + numInicial < 10)
                                {
                                    novoTexto += "0" + $"{numInicial + i}";
                                }
                                else
                                {
                                    novoTexto += $"{numInicial + i}";
                                }
                            }
                            else if (numFinal > 100 && formVal.frmZeros == true)
                            {
                                if (i + numInicial < 10)
                                {
                                    novoTexto += "00" + $"{numInicial + i}";
                                }
                                else if (i + numInicial < 100)
                                {
                                    novoTexto += "0" + $"{numInicial + i}";
                                }
                                else
                                {
                                    novoTexto += $"{numInicial + i}";
                                }
                            }
                            else
                            {
                                novoTexto += $"{numInicial + i}";
                            }

                            // Caso tenha sido definido um sufixo, este é adicionado ao texto para substituição
                            if (formVal.frmSufixo != "")
                            {
                                novoTexto += formVal.frmSufixo;
                            }

                            // Define os valores de dimensões do objeto inicial para definir o posicionamento
                            corelApp.ActiveSelection.GetSize(out offX1, out offY1);

                            // Substitui o texto da seleção de acordo com o texto definido anteriormente
                            corelApp.ActiveSelection.Shapes.FindShape(Type: corel.cdrShapeType.cdrTextShape).Text.Replace(OldText: txtExistente,
                                                                                                                          NewText: novoTexto,
                                                                                                                          CaseSensitive: true);

                            // Define os valores de dimensões do objeto atualizado para definir o posicionamento
                            corelApp.ActiveSelection.GetSize(out offX2, out offY2);

                            // Calcula o posicionamento do novo objeto numerado
                            double offDup = (offX1 / 2) + (offX2 / 2) + 0.5;

                            // Duplica a seleção de acordo com o posicionamento definido anteriormente, multiplicando de acordo com o tipo de objeto
                            corelApp.ActiveSelection.Duplicate(i * offDup, 0);
                        }

                        // Remove a seleção inicial para evitar a existência de sobreposição de objetos
                        corelApp.ActiveSelection.Delete();
                    }
                }
            }
        }

        // Exportar Arquivo em Curvas - Dea! Design
        private void expCurvas(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Inicia o formulário para coleta das variáveis
            FormExpCurvas formPasta = new FormExpCurvas();

            System.Windows.Forms.Application.Run(formPasta);

            // Inicia a ação do formulário após o usuário apertar 'Iniciar'
            if (formPasta.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                // Verifica o total de páginas do documento para fazer a transformação
                int pageCount = corelApp.ActiveDocument.Pages.Count;

                // Altera a página ativa para a página inicial do Documento (evita que a iteração se perca na contagem)
                corelApp.ActiveDocument.Pages[1].Activate();

                // Inicia a iteração das páginas, transformando em curvas os objetos da página ativa
                foreach (corel.Page p in corelApp.ActiveDocument.Pages)
                {
                    // Cria instância do número da página
                    int pNumber = p.Index;

                    // Ativa a página indicada pela instância pNumber
                    corelApp.ActiveDocument.Pages[pNumber].Activate();

                    // Verifica se na página existem símbolos, caso sim, converte todos em objetos
                    foreach (corel.Shape s in corelApp.ActivePage.SelectableShapes)
                    {
                        if (s.Type == corel.cdrShapeType.cdrSymbolShape)
                        {
                            s.AddToSelection();
                            s.Symbol.RevertToShapes();
                            corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();
                        }
                    }

                    // Desagrupa todos os itens da página para evitar que objetos não sejam convertidos
                    corelApp.ActivePage.SelectableShapes.All().AddToSelection();
                    corelApp.ActiveSelection.UngroupAll();

                    // Remove qualquer item de seleção para fazer as iterações sem distorções
                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();

                    // Verifica se na página existem imagens e curvas fechadas e inclui o índice em uma lista para remoção da seleção
                    List<int> shapeIndex = new List<int>();
                    foreach (corel.Shape s in corelApp.ActivePage.SelectableShapes)
                    {
                        if (s.Type == corel.cdrShapeType.cdrBitmapShape)
                        {
                            shapeIndex.Add(corelApp.ActivePage.SelectableShapes.All().IndexOf(s));
                        }
                    }
                    int[] selIndexes = shapeIndex.ToArray();

                    // Seleciona todos os objetos presentes na página, com exceção das imagens e PowerClips
                    corelApp.ActivePage.SelectableShapes.AllExcluding(selIndexes).AddToSelection();

                    // Transforma os objetos em curvas e limpa a seleção
                    corelApp.ActiveSelection.ConvertToCurves();
                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();

                    //Verifica quais são os objetos em linhas abertas e transforma os contornos em curvas
                    foreach (corel.Shape s in corelApp.ActivePage.SelectableShapes)
                    {
                        if (s.Type != corel.cdrShapeType.cdrCurveShape)
                        {
                            continue;
                        }
                        else if (s.Curve.Closed == true || s.Outline.Type == corel.cdrOutlineType.cdrNoOutline || s.Curve.Length == 0)
                        {
                            continue;
                        }
                        else
                        {
                            s.Outline.ConvertToObject();
                        }
                    }

                    // Faz a mesma interação de converter em curvas mas para objetos dentro de Power Clips (se houver)
                    foreach (corel.Shape s in corelApp.ActivePage.SelectableShapes)
                    {
                        if (s.PowerClip is corel.PowerClip)
                        {
                            // Entra no modo de edição do PowerClip
                            s.PowerClip.EnterEditMode();

                            // Verifica se no PowerClip existem símbolos, caso sim, converte todos em objetos
                            foreach (corel.Shape smb in corelApp.ActiveLayer.SelectableShapes.All())
                            {
                                if (smb.Type == corel.cdrShapeType.cdrSymbolShape)
                                {
                                    smb.AddToSelection();
                                    smb.Symbol.RevertToShapes();
                                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();
                                }
                            }

                            // Converte em curvas os conteúdos do PowerClip
                            foreach (corel.Shape spc in corelApp.ActiveLayer.SelectableShapes.All())
                            {
                                // Verifica se o objeto é uma imagem, do contrário adiciona à seleção e transforma em curvas
                                if (spc.Type == corel.cdrShapeType.cdrBitmapShape)
                                {
                                    continue;
                                }
                                else
                                {
                                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();
                                    spc.AddToSelection();
                                    spc.ConvertToCurves();
                                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();
                                }
                            }

                            // Converte em objeto os contornos
                            foreach (corel.Shape spc in corelApp.ActiveLayer.SelectableShapes.All())
                            {
                                // Verifica se o objeto é uma imagem ou uma curva fechada ou sem contorno, do contrário adiciona à seleção e transforma em objeto
                                if (spc.Type != corel.cdrShapeType.cdrCurveShape)
                                {
                                    continue;
                                }
                                else if (spc.Curve.Closed == true || spc.Outline.Type == corel.cdrOutlineType.cdrNoOutline)
                                {
                                    continue;
                                }
                                else
                                {
                                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();
                                    spc.AddToSelection();
                                    spc.Outline.ConvertToObject();
                                    corelApp.ActivePage.SelectableShapes.All().RemoveFromSelection();
                                }
                            }

                            // Sai do modo de edição do PowerClip
                            s.PowerClip.LeaveEditMode();
                        }
                    }
                }

                // Salva o arquivo Corel com final do nome do arquivo _Curvas
                corel.StructSaveAsOptions cdrSave = new corel.StructSaveAsOptions();
                cdrSave.Overwrite = true;

                string cdrFilename = corelApp.ActiveDocument.Title.Substring(0, corelApp.ActiveDocument.Title.Length - 4);

                string newCdrFilename = formPasta.caminhoPastaTxt + "\\" + cdrFilename + "_Curvas" + ".cdr";

                corelApp.ActiveDocument.SaveAs(newCdrFilename, Options: cdrSave);

                // Salva o arquivo PDF com final do nome do arquivo _Curvas
                string newPdfFilename = formPasta.caminhoPastaTxt + "\\" + cdrFilename + ".pdf";

                corelApp.ActiveDocument.PublishToPDF(newPdfFilename);

                // Finaliza com mensagem de êxito
                System.Windows.MessageBox.Show("Caderno exportado em curvas com sucesso!", "Atenção!");
            }
        }

        // Cotas Inteiras - Dea! Design
        private void cotaInteira(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Verifica se o usuário está selecionando objetos
            if (corelApp.ActiveSelectionRange.Count == 0 || corelApp.ActiveSelectionRange.Count > 1)
            {
                System.Windows.MessageBox.Show("Selecione apenas um objeto e/ou grupo para cotar.", "Atenção!");
            }
            else
            {
                // Define o dashed style das linhas de cota
                corel.OutlineStyle dsh = new corel.OutlineStyle();
                dsh.DashCount = 2;
                dsh.DashLength[1] = 6;
                dsh.GapLength[1] = 6;

                // Define a escala de unidades do documento
                corelApp.ActiveDocument.Unit = corelApp.ActiveDocument.Rulers.HUnits;
                corel.cdrUnit unidOriginal = corelApp.ActiveDocument.Unit;

                // Define o valor da dimensão horizontal
                double dimHor = corelApp.ActiveSelection.SizeWidth;

                // Define o valor da dimensão vertical
                double dimVer = corelApp.ActiveSelection.SizeHeight;

                // Torna as unidades do documento em cm para posicionamento e peso de linha
                if (unidOriginal != corel.cdrUnit.cdrCentimeter)
                {
                    corelApp.ActiveDocument.Unit = corel.cdrUnit.cdrCentimeter;
                }

                // Define as coordenadas de posicionamento da linha de cota horizontal
                double xh1 = corelApp.ActiveSelection.LeftX;
                double xh2 = corelApp.ActiveSelection.RightX;
                double yh1 = corelApp.ActiveSelection.TopY + 0.5;
                double yh2 = corelApp.ActiveSelection.TopY;

                // Define as coordenadas de posicionamento da linha de cota vertical
                double xv1 = corelApp.ActiveSelection.LeftX;
                double xv2 = corelApp.ActiveSelection.LeftX - 0.5;
                double yv1 = corelApp.ActiveSelection.TopY;
                double yv2 = corelApp.ActiveSelection.BottomY;

                // Cria as linhas da cota horizontal
                corel.Shape linCon1 = corelApp.ActiveLayer.CreateLineSegment(xh1, yh2, xh1, yh1);
                linCon1.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon1.Outline.Width = 0.01;
                linCon1.Outline.Style = dsh;

                corel.Shape linCon2 = corelApp.ActiveLayer.CreateLineSegment(xh2, yh2, xh2, yh1);
                linCon2.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon2.Outline.Width = 0.01;
                linCon2.Outline.Style = dsh;

                corel.Shape linCota = corelApp.ActiveLayer.CreateLineSegment(xh1, yh1, xh2, yh1);
                linCota.Outline.Color.CMYKAssign(0, 0, 0, 100);
                linCota.Outline.StartArrow = corelApp.ArrowHeads[59];
                linCota.Outline.EndArrow = corelApp.ArrowHeads[59];
                linCota.Outline.Width = 0.01;

                // Cria o valor da cota baseado nas dimensões em escala do objeto
                double cotaH = dimHor * corelApp.ActiveDocument.WorldScale;
                cotaH = Math.Round(cotaH, 1);

                corel.Shape txtCotaH = corelApp.ActiveLayer.CreateArtisticText(Left: 0,
                                                                               Bottom: 0,
                                                                               Text: cotaH.ToString(),
                                                                               Font: "Arial",
                                                                               Size: 8,
                                                                               Alignment: corel.cdrAlignment.cdrCenterAlignment);

                txtCotaH.AlignToShape(corel.cdrAlignType.cdrAlignBottom, linCota);
                txtCotaH.AlignToShape(corel.cdrAlignType.cdrAlignHCenter, linCota);
                txtCotaH.Move(0, 0.15);

                // Cria as linhas da cota vertical
                corel.Shape linCon3 = corelApp.ActiveLayer.CreateLineSegment(xv1, yv1, xv2, yv1);
                linCon3.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon3.Outline.Width = 0.01;
                linCon3.Outline.Style = dsh;

                corel.Shape linCon4 = corelApp.ActiveLayer.CreateLineSegment(xv1, yv2, xv2, yv2);
                linCon4.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon4.Outline.Width = 0.01;
                linCon4.Outline.Style = dsh;

                corel.Shape linCota2 = corelApp.ActiveLayer.CreateLineSegment(xv2, yv1, xv2, yv2);
                linCota2.Outline.Color.CMYKAssign(0, 0, 0, 100);
                linCota2.Outline.StartArrow = corelApp.ArrowHeads[59];
                linCota2.Outline.EndArrow = corelApp.ArrowHeads[59];
                linCota2.Outline.Width = 0.01;

                // Cria o valor da cota baseado nas dimensões em escala do objeto
                double cotaV = dimVer * corelApp.ActiveDocument.WorldScale;
                cotaV = Math.Round(cotaV, 1);

                corel.Shape txtCotaV = corelApp.ActiveLayer.CreateArtisticText(Left: 0,
                                                                               Bottom: 0,
                                                                               Text: cotaV.ToString(),
                                                                               Font: "Arial",
                                                                               Size: 8,
                                                                               Alignment: corel.cdrAlignment.cdrCenterAlignment);

                txtCotaV.Rotate(90);
                txtCotaV.AlignToShape(corel.cdrAlignType.cdrAlignRight, linCota2);
                txtCotaV.AlignToShape(corel.cdrAlignType.cdrAlignVCenter, linCota2);
                txtCotaV.Move(-0.15, 0);

                corelApp.ActiveDocument.ClearSelection();

                corelApp.ActiveDocument.AddToSelection(linCon1, linCon2, linCota);
                corelApp.ActiveSelection.Group();
                corelApp.ActiveDocument.ClearSelection();

                corelApp.ActiveDocument.AddToSelection(linCon3, linCon4, linCota2);
                corelApp.ActiveSelection.Group();
                corelApp.ActiveDocument.ClearSelection();

                // Retoma as unidades originais pro caso do arquivo não estar em cm
                if (unidOriginal != corel.cdrUnit.cdrCentimeter)
                {
                    corelApp.ActiveDocument.Unit = unidOriginal;
                }
            }
        }

        // Cotas Decimais - Dea! Design
        private void cotaDecimal(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Verifica se o usuário está selecionando objetos
            if (corelApp.ActiveSelectionRange.Count == 0 || corelApp.ActiveSelectionRange.Count > 1)
            {
                System.Windows.MessageBox.Show("Selecione apenas um objeto e/ou grupo para cotar.", "Atenção!");
            }
            else
            {
                // Define o dashed style das linhas de cota
                corel.OutlineStyle dsh = new corel.OutlineStyle();
                dsh.DashCount = 2;
                dsh.DashLength[1] = 6;
                dsh.GapLength[1] = 6;

                // Define a escala de unidades do documento
                corelApp.ActiveDocument.Unit = corelApp.ActiveDocument.Rulers.HUnits;
                corel.cdrUnit unidOriginal = corelApp.ActiveDocument.Unit;

                // Define o valor da dimensão horizontal
                double dimHor = corelApp.ActiveSelection.SizeWidth;

                // Define o valor da dimensão vertical
                double dimVer = corelApp.ActiveSelection.SizeHeight;

                // Torna as unidades do documento em cm para posicionamento e peso de linha
                if (unidOriginal != corel.cdrUnit.cdrCentimeter)
                {
                    corelApp.ActiveDocument.Unit = corel.cdrUnit.cdrCentimeter;
                }

                // Define as coordenadas de posicionamento da linha de cota horizontal
                double xh1 = corelApp.ActiveSelection.LeftX;
                double xh2 = corelApp.ActiveSelection.RightX;
                double yh1 = corelApp.ActiveSelection.TopY + 0.5;
                double yh2 = corelApp.ActiveSelection.TopY;

                // Define as coordenadas de posicionamento da linha de cota vertical
                double xv1 = corelApp.ActiveSelection.LeftX;
                double xv2 = corelApp.ActiveSelection.LeftX - 0.5;
                double yv1 = corelApp.ActiveSelection.TopY;
                double yv2 = corelApp.ActiveSelection.BottomY;

                // Cria as linhas da cota horizontal
                corel.Shape linCon1 = corelApp.ActiveLayer.CreateLineSegment(xh1, yh2, xh1, yh1);
                linCon1.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon1.Outline.Width = 0.01;
                linCon1.Outline.Style = dsh;

                corel.Shape linCon2 = corelApp.ActiveLayer.CreateLineSegment(xh2, yh2, xh2, yh1);
                linCon2.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon2.Outline.Width = 0.01;
                linCon2.Outline.Style = dsh;

                corel.Shape linCota = corelApp.ActiveLayer.CreateLineSegment(xh1, yh1, xh2, yh1);
                linCota.Outline.Color.CMYKAssign(0, 0, 0, 100);
                linCota.Outline.StartArrow = corelApp.ArrowHeads[59];
                linCota.Outline.EndArrow = corelApp.ArrowHeads[59];
                linCota.Outline.Width = 0.01;

                // Cria o valor da cota baseado nas dimensões em escala do objeto
                double cotaH = dimHor * corelApp.ActiveDocument.WorldScale;
                cotaH = Math.Round(cotaH, 1);

                corel.Shape txtCotaH = corelApp.ActiveLayer.CreateArtisticText(Left: 0,
                                                                               Bottom: 0,
                                                                               Text: cotaH.ToString("0.0"),
                                                                               Font: "Arial",
                                                                               Size: 8,
                                                                               Alignment: corel.cdrAlignment.cdrCenterAlignment);

                txtCotaH.AlignToShape(corel.cdrAlignType.cdrAlignBottom, linCota);
                txtCotaH.AlignToShape(corel.cdrAlignType.cdrAlignHCenter, linCota);
                txtCotaH.Move(0, 0.15);

                // Cria as linhas da cota vertical
                corel.Shape linCon3 = corelApp.ActiveLayer.CreateLineSegment(xv1, yv1, xv2, yv1);
                linCon3.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon3.Outline.Width = 0.01;
                linCon3.Outline.Style = dsh;

                corel.Shape linCon4 = corelApp.ActiveLayer.CreateLineSegment(xv1, yv2, xv2, yv2);
                linCon4.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon4.Outline.Width = 0.01;
                linCon4.Outline.Style = dsh;

                corel.Shape linCota2 = corelApp.ActiveLayer.CreateLineSegment(xv2, yv1, xv2, yv2);
                linCota2.Outline.Color.CMYKAssign(0, 0, 0, 100);
                linCota2.Outline.StartArrow = corelApp.ArrowHeads[59];
                linCota2.Outline.EndArrow = corelApp.ArrowHeads[59];
                linCota2.Outline.Width = 0.01;

                // Cria o valor da cota baseado nas dimensões em escala do objeto
                double cotaV = dimVer * corelApp.ActiveDocument.WorldScale;
                cotaV = Math.Round(cotaV, 1);

                corel.Shape txtCotaV = corelApp.ActiveLayer.CreateArtisticText(Left: 0,
                                                                               Bottom: 0,
                                                                               Text: cotaV.ToString("0.0"),
                                                                               Font: "Arial",
                                                                               Size: 8,
                                                                               Alignment: corel.cdrAlignment.cdrCenterAlignment);

                txtCotaV.Rotate(90);
                txtCotaV.AlignToShape(corel.cdrAlignType.cdrAlignRight, linCota2);
                txtCotaV.AlignToShape(corel.cdrAlignType.cdrAlignVCenter, linCota2);
                txtCotaV.Move(-0.15, 0);

                corelApp.ActiveDocument.ClearSelection();

                corelApp.ActiveDocument.AddToSelection(linCon1, linCon2, linCota);
                corelApp.ActiveSelection.Group();
                corelApp.ActiveDocument.ClearSelection();

                corelApp.ActiveDocument.AddToSelection(linCon3, linCon4, linCota2);
                corelApp.ActiveSelection.Group();
                corelApp.ActiveDocument.ClearSelection();

                // Retoma as unidades originais pro caso do arquivo não estar em cm
                if (unidOriginal != corel.cdrUnit.cdrCentimeter)
                {
                    corelApp.ActiveDocument.Unit = unidOriginal;
                }
            }
        }

        // Cotas Alinhadas - Dea! Design
        private void cotaAlinhada(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Verifica se o usuário está selecionando objetos
            if (corelApp.ActiveSelectionRange.Count == 0 || corelApp.ActiveSelectionRange.Count > 1)
            {
                System.Windows.MessageBox.Show("Selecione apenas um objeto e/ou grupo para cotar.", "Atenção!");
            }
            else
            {
                // Define o dashed style das linhas de cota
                corel.OutlineStyle dsh = new corel.OutlineStyle();
                dsh.DashCount = 2;
                dsh.DashLength[1] = 6;
                dsh.GapLength[1] = 6;

                // Define a escala de unidades do documento
                corelApp.ActiveDocument.Unit = corelApp.ActiveDocument.Rulers.HUnits;
                corel.cdrUnit unidOriginal = corelApp.ActiveDocument.Unit;

                // Pede para que o usuário clique no primeiro ponto da cota
                double x1 = 0.0;
                double y1 = 0.0;
                int aux = 0;
                corelApp.ActiveDocument.GetUserClick(out x1, out y1, out aux, 10, true, corel.cdrCursorShape.cdrCursorPickOvertarget);

                // Pede para que o usuário clique no segundo ponto da cota
                double x2 = 0.0;
                double y2 = 0.0;
                corelApp.ActiveDocument.GetUserClick(out x2, out y2, out aux, 10, true, corel.cdrCursorShape.cdrCursorPickOvertarget);

                // Calcula a medida da cota com base na escala do documento
                double dist = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2));

                // Converte a distância de posicionamento da cota de 0.5 cm para unidade da régua
                double conv = corelApp.ConvertUnits(0.5, corel.cdrUnit.cdrCentimeter, corelApp.ActiveDocument.Unit);

                // Define as coordenadas de posicionamento da linha de cota antes da rotação
                double xh1 = x1;
                double xh2 = x1 + dist;
                double yh1 = y1 + conv;
                double yh2 = y1;

                // Cria as linhas da cota antes da rotação
                corel.Shape linCon1 = corelApp.ActiveLayer.CreateLineSegment(xh1, yh2, xh1, yh1);
                corel.Shape linCon2 = corelApp.ActiveLayer.CreateLineSegment(xh2, yh2, xh2, yh1);
                corel.Shape linCota = corelApp.ActiveLayer.CreateLineSegment(xh1, yh1, xh2, yh1);

                // Calcula o ângulo da cota em relação ao plano (sempre considera que a escolha dos pontos foi em sentido horário)
                double rx = x2 - x1;
                double ry = y2 - y1;

                double ang = Math.Atan2(ry, rx);

                // Torna as unidades do documento em cm para peso de linha
                if (unidOriginal != corel.cdrUnit.cdrCentimeter)
                {
                    corelApp.ActiveDocument.Unit = corel.cdrUnit.cdrCentimeter;
                }

                // Ajusta aparência da cota
                linCon1.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon1.Outline.Width = 0.01;
                linCon1.Outline.Style = dsh;

                linCon2.Outline.Color.CMYKAssign(0, 0, 0, 50);
                linCon2.Outline.Width = 0.01;
                linCon2.Outline.Style = dsh;

                linCota.Outline.Color.CMYKAssign(0, 0, 0, 100);
                linCota.Outline.StartArrow = corelApp.ArrowHeads[59];
                linCota.Outline.EndArrow = corelApp.ArrowHeads[59];
                linCota.Outline.Width = 0.01;

                // Arredonda o valor e cria o texto da cota
                double cota = Math.Round(dist * corelApp.ActiveDocument.WorldScale, 1);

                corel.Shape txtCotaH = corelApp.ActiveLayer.CreateArtisticText(Left: 0,
                                                                               Bottom: 0,
                                                                               Text: cota.ToString(),
                                                                               Font: "Arial",
                                                                               Size: 8,
                                                                               Alignment: corel.cdrAlignment.cdrCenterAlignment);

                txtCotaH.AlignToShape(corel.cdrAlignType.cdrAlignBottom, linCota);
                txtCotaH.AlignToShape(corel.cdrAlignType.cdrAlignHCenter, linCota);
                txtCotaH.Move(0, 0.15);

                // Retoma as unidades originais pro caso do arquivo não estar em cm
                if (unidOriginal != corel.cdrUnit.cdrCentimeter)
                {
                    corelApp.ActiveDocument.Unit = unidOriginal;
                }

                //Agrupa e rotaciona o texto
                corelApp.ActiveDocument.ClearSelection();

                corelApp.ActiveDocument.AddToSelection(linCon1, linCon2, linCota, txtCotaH);
                corelApp.ActiveSelection.Group();
                corelApp.ActiveSelection.RotationCenterX = x1;
                corelApp.ActiveSelection.RotationCenterY = y1;
                corelApp.ActiveSelection.Rotate(ang * 180 / Math.PI);

                corelApp.ActiveDocument.ClearSelection();
            }
        }

        // Excel para Placas - Dea! Design
        private void excelPlacas(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Verifica se o usuário está selecionando objetos
            if (corelApp.ActiveSelectionRange.Count == 0)
            {
                System.Windows.MessageBox.Show("Selecione um objeto e/ou grupo para iniciar a a inserção de conteúdo, já contendo os textos de base formatados.", "Atenção!");
            }
            else
            {
                // Verifica se os objetos selecionados incluem textos
                if (corelApp.ActiveSelection.Shapes.FindShapes(Type: corel.cdrShapeType.cdrTextShape).Count < 1)
                {
                    System.Windows.MessageBox.Show("Selecione um objeto/grupo que inclua ao menos um texto para inserção de conteúdo.", "Atenção!");
                }
                else
                {
                    // Inicia o formulário para coleta das variáveis
                    FormAbrirExcel formExcel = new FormAbrirExcel();

                    System.Windows.Forms.Application.Run(formExcel);

                    // Inicia a ação do formulário após o usuário apertar 'Iniciar'
                    if (formExcel.DialogResult == System.Windows.Forms.DialogResult.OK)
                    {
                        // Verifica se o arquivo escolhido é do formato Excel (.xls ou .xlsx)
                        string excelName = formExcel.caminhoExcelTxt;
                        FileInfo excelInfo = new FileInfo(excelName);
                        if (excelInfo.Extension != ".xlsx")
                        {
                            System.Windows.MessageBox.Show("Arquivo selecionado não é de formato Excel! (.xls ou .xlsx)", "Atenção!");
                            return;
                        }

                        // Verifica se a inserção das colunas foi feita corretamente
                        foreach (string s in formExcel.colunasExcel)
                        {
                            if (s.Contains(" ") || s.Contains("#") || s.Contains("_") || s.Contains("-") || s.Contains(".") || s.Contains(","))
                            {
                                System.Windows.MessageBox.Show("Confirme se a inserção das colunas foi feita corretamente, com separação de ponto vírgula sem espaço entre o nome das colunas (ex.: A;B;C).", "Atenção!");
                                return;
                            }
                        }

                        // Inicializa o Excel
                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelName);

                        // Verifica se a planilha indicada existe no arquivo, caso não exista, cancela operação. Se existir, define o alcance de colunas e linhas preenchidas.
                        List<string> planilhas = new List<string>();
                        foreach (Excel.Worksheet w in excelWorkbook.Worksheets)
                        {
                            planilhas.Add(w.Name);
                        }
                        if (planilhas.Contains(formExcel.planilhaExcel) == false)
                        {
                            System.Windows.MessageBox.Show("Planilha inserida no formulário não existe no arquivo!", "Atenção!");
                            excelWorkbook.Close(0);
                            excelApp.Quit();
                            return;
                        }
                        Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets[formExcel.planilhaExcel];
                        int linhasTotal = excelWorksheet.UsedRange.Rows.Count;
                        int colunasTotal = excelWorksheet.UsedRange.Columns.Count;

                        // Converte as colunas do formulário em índices (limite A-Z) e armazena o título da coluna em uma lista
                        Dictionary<string, int> indexCol = new Dictionary<string, int>();
                        foreach (string s in formExcel.colunasExcel)
                        {
                            int index = 0;

                            if (s.Length > 1)
                            {
                                index = ((char.ToUpper(s[0]) - 64) * 26) + (char.ToUpper(s[1]) - 64);
                            }
                            else if (s.Length > 2)
                            {
                                System.Windows.MessageBox.Show("Essa função está limitada a trabalhar até a coluna ZZ!", "Atenção!");
                                excelWorkbook.Close(0);
                                excelApp.Quit();
                                return;
                            }
                            else
                            {
                                index = char.ToUpper(s[0]) - 64;
                            }

                            indexCol.Add(excelWorksheet.Cells[1, index].Value.ToString(), index);
                        }

                        // Verifica se as colunas indicadas possuem texto correspondente no grupo escolhido
                        List<string> textosPlaca = new List<string>();

                        corelApp.ActiveSelection.UngroupAll();
                        foreach (corel.Shape s in corelApp.ActiveSelection.Shapes)
                        {
                            if (s.Type == corel.cdrShapeType.cdrTextShape)
                            {
                                textosPlaca.Add(s.Text.Contents);
                            }
                        }
                        corelApp.ActiveSelection.Group();
                        foreach (string s in indexCol.Keys)
                        {
                            if (textosPlaca.Contains(s) == false)
                            {
                                System.Windows.MessageBox.Show($"Coluna {s} não possui correspondência com nenhum objeto de texto do grupo selecionado.", "Atenção!");
                                return;
                            }
                        }

                        // Armazena em um dicionário os valores das colunas utilizadas
                        Dictionary<string, List<string>> valoresExcel = new Dictionary<string, List<string>>();
                        foreach (string s in indexCol.Keys)
                        {
                            valoresExcel.Add(s, new List<string>());

                            int tempCount = 2;
                            while (tempCount < linhasTotal + 1)
                            {
                                if (excelWorksheet.Cells[tempCount, indexCol[s]].Value2 != null)
                                {
                                    valoresExcel[s].Add(excelWorksheet.Cells[tempCount, indexCol[s]].Value2.ToString());
                                    tempCount++;
                                }
                                else
                                {
                                    valoresExcel[s].Add(excelWorksheet.Cells[tempCount, indexCol[s]].Value2);
                                    tempCount++;
                                }                                                              
                            }
                        }

                        // Finaliza os processos do Excel
                        excelWorkbook.Close(0);
                        excelApp.Quit();

                        // Declara e aloca valores na lista de IDs dos objetos selecionados
                        List<int> idsOriginais = new List<int>();
                        foreach (corel.Shape s in corelApp.ActiveSelection.Shapes)
                        {
                            idsOriginais.Add(s.StaticID);
                        }

                        // Declara a lista de IDs dos novos objetos, para posterior interação com as novas placas criadas
                        Dictionary<int, int> ids = new Dictionary<int, int>();

                        // Copia a seleção de objetos e altera o posicionamento da cópia
                        for (int i = 0; i < linhasTotal - 1; i++)
                        {
                            // Limpa a seleção de objetos para garantir que nada além da placa original esteja selecionada
                            corelApp.ActiveDocument.ClearSelection();

                            // Adiciona os objetos da seleção original à seleção de cópia
                            foreach (int id in idsOriginais)
                            {
                                corelApp.ActiveLayer.FindShape(StaticID: id).AddToSelection();
                            }
                        
                            // Declara as variáveis de posicionamento da nova cópia
                            double offX1 = 0.0;
                            double offY1 = 0.0;

                            // Define os valores de dimensões do objeto inicial para definir o posicionamento
                            corelApp.ActiveSelection.GetSize(out offX1, out offY1);

                            // Duplica a seleção de acordo com o posicionamento definido
                            double offX2 = ((i % 5) * (offX1 + 0.5)) + offX1 + 0.5;
                            double offY2 = (Math.Floor(i / 5d) * offY1) + (Math.Floor(i / 5d) * 1.5);
                            corelApp.ActiveSelection.Copy();
                            corel.Shape s = corelApp.ActiveLayer.Paste();
                            s.Move(offX2, -offY2);

                            // Adiciona o ID da nova placa à lista de IDs
                            ids.Add(i + 1, s.StaticID);
                        }

                        // Limpa a seleção de objetos para iniciar a substituição de textos
                        corelApp.ActiveDocument.ClearSelection();

                        // Inicia a substituição dos textos de acordo com as colunas inseridas pelo usuário
                        for (int i = 1; i < linhasTotal; i++)
                        {
                            corelApp.ActiveLayer.FindShape(StaticID: ids[i]).AddToSelection();
                            corelApp.ActiveSelection.Ungroup();

                            foreach (corel.Shape s in corelApp.ActiveSelection.Shapes)
                            {
                                if (s.Type == corel.cdrShapeType.cdrTextShape && indexCol.Keys.ToList<string>().Contains(s.Text.Contents))
                                {
                                    if (valoresExcel[s.Text.Contents][i - 1] == null)
                                    {
                                        s.Delete();
                                    }
                                    else
                                    {
                                        string novoTexto = valoresExcel[s.Text.Contents][i - 1];
                                        s.Text.Replace(OldText: s.Text.Contents,
                                                       NewText: novoTexto,
                                                       CaseSensitive: true);
                                    }                                                              
                                }
                            }
                            corelApp.ActiveSelection.Group();
                            corelApp.ActiveDocument.ClearSelection();
                        }
                    }
                }
            }
        }

        // Exportar Arquivo em Imagens - Dea! Design
        private void expPNG(object sender, RoutedEventArgs e)
        {
            // Verifica se o Corel está com arquivos ativos, caso contrário retorna um erro
            if (corelApp.Documents.Count == 0)
            {
                System.Windows.MessageBox.Show("Para que esta ação ocorra algum documento precisa estar aberto no CorelDRAW.", "Atenção!");
                return;
            }

            // Inicia o formulário para coleta das variáveis
            FormExpPNG formPasta = new FormExpPNG();

            System.Windows.Forms.Application.Run(formPasta);

            // Inicia a ação do formulário após o usuário apertar 'Iniciar'
            if (formPasta.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                // Verifica o total de páginas do documento para fazer a transformação
                int pageCount = corelApp.ActiveDocument.Pages.Count;

                // Altera a página ativa para a página inicial do Documento (evita que a iteração se perca na contagem)
                corelApp.ActiveDocument.Pages[1].Activate();

                // Inicia a iteração das páginas, transformando em curvas os objetos da página ativa
                foreach (corel.Page p in corelApp.ActiveDocument.Pages)
                {
                    // Cria instância do número da página
                    int pNumber = p.Index;

                    // Ativa a página indicada pela instância pNumber
                    corelApp.ActiveDocument.Pages[pNumber].Activate();

                    // Define o nome do arquivo e o caminho para salvar a imagem
                    string nomePNG = corelApp.ActiveDocument.Title.Substring(0, corelApp.ActiveDocument.Title.Length - 4) + "_" + pNumber;
                    string nomeArquivo = formPasta.caminhoPastaTxt + "\\" + nomePNG + ".png";

                    // Define as opções de exportação e salva a imagem no caminho indicado
                    corel.StructExportOptions optSalvar = new corel.StructExportOptions();

                    optSalvar.ImageType = corel.cdrImageType.cdrRGBColorImage;
                    optSalvar.Transparent = formPasta.alphaEscolha;
                    optSalvar.MaintainAspect = true;
                    optSalvar.ResolutionX = formPasta.resolEscolha;
                    optSalvar.ResolutionY = formPasta.resolEscolha;
                    optSalvar.SizeX = Convert.ToInt32(p.SizeWidth * optSalvar.ResolutionX);
                    optSalvar.SizeY = Convert.ToInt32(p.SizeHeight * optSalvar.ResolutionY);

                    corelApp.ActiveDocument.Export(FileName: nomeArquivo,
                                                   Filter: corel.cdrFilter.cdrPNG,
                                                   Range: corel.cdrExportRange.cdrCurrentPage,
                                                   Options: optSalvar);
                }
            }
        }
    }
}
