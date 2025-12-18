using Ms.Skills.POS.Services;
using Primavera.Extensibility.BusinessEntities;
using Primavera.Extensibility.BusinessEntities.ExtensibilityService.EventArgs;
using Primavera.Extensibility.Integration;
using Primavera.Extensibility.POS.Editors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.LinkLabel;

namespace Ms.Skills.POS.POS
{
    public class UiEditorVendas : EditorVendas
    {
        public override void TeclaPressionada(int KeyCode, int Shift, ExtensibilityEventArgs e)
        {
            base.TeclaPressionada(KeyCode, Shift, e);

            try
            {
                PriEngine.Engine = BSO;
                PriEngine.Platform = PSO;

                var engine = new MainSrv();

                if (Shift == 2 && (KeyCode == 68 || KeyCode == 100))
                {
                    string tipodoc = engine.GetParameter("POS_TipoDoc");

                    string query = $@"select cd.Id,Documento, Serie, TipoDoc,NumDoc, Nome, NumContribuinte, Morada, TotalDocumento from cabecdoc cd
	                                    inner join CabecDocStatus cds on cds.IdCabecDoc = cd.Id and cds.Estado = 'P' 
		                                    and cds.Fechado = 0 and cds.Anulado = 0
		                                    order by [Data] desc
                                    where cd.tipodoc = '{tipodoc}'
                                    ";
                    string coluna = "Id";
                    string tabela = "Lista de Contas Bancarias";
                    string resultado = engine.F4Event(query, coluna, tabela).ToString();

                    if(resultado.Length> 0)
                    {
                        var _doc = PriEngine.Engine.Vendas.Documentos.EditaID(resultado);

                        if (_doc.Linhas.NumItens > 0) 
                        {
                            this.DocumentoVenda.Linhas.RemoveTodos();

                            BSO.Vendas.Documentos.AdicionaLinhaEspecial(_doc, BasBE100.BasBETiposGcp.vdTipoLinhaEspecial.vdLinha_Comentario, 0,
                                _doc.Documento);

                            BSO.Vendas.Documentos.AdicionaLinhaEspecial(_doc, BasBE100.BasBETiposGcp.vdTipoLinhaEspecial.vdLinha_Comentario, 0,
                                "");

                            foreach(var item in _doc.Linhas.Where(p=>p.Artigo.Length>0))
                            {
                                if (BSO.Base.Artigos.Existe(item.Artigo))
                                {
                                    string armazem = item.Armazem;
                                    string localizacao = item.Armazem;
                                    string tipoLinha = item.TipoLinha;
                                    double qnt = item.Quantidade;

                                    BSO.Vendas.Documentos.AdicionaLinha(_doc,
                                        item.Artigo,
                                        ref qnt,
                                        ref armazem,
                                        ref localizacao,
                                        item.PrecUnit,
                                        item.Desconto1,
                                        item.Lote
                                    );

                                    var numLinha = _doc.Linhas.NumItens;

                                    _doc.Linhas.GetEdita(numLinha).IDLinhaOriginal = item.IDLinhaOriginal;

                                    _doc.Linhas.GetEdita(numLinha).Armazem = item.Armazem;
                                    _doc.Linhas.GetEdita(numLinha).Localizacao = item.Armazem;

                                    _doc.Linhas.GetEdita(numLinha).Lote = item.Lote;
                                    _doc.Linhas.GetEdita(numLinha).Quantidade = item.Quantidade;
                                    _doc.Linhas.GetEdita(numLinha).PrecUnit = item.PrecUnit;

                                    _doc.Linhas.GetEdita(numLinha).DataEntrega = DateTime.Now;
                                    _doc.Linhas.GetEdita(numLinha).DataStock = DateTime.Now;
                                    _doc.Linhas.GetEdita(numLinha).FactorConv = Math.Round(item.FactorConv, 10);

                                    _doc.Linhas.GetEdita(numLinha).Unidade = item.Unidade;
                                    _doc.Linhas.GetEdita(numLinha).MovStock = item.MovStock?.Length > 0 ? item.MovStock : "N";
                                    _doc.Linhas.GetEdita(numLinha).ArredFConv = 0;
                                    _doc.Linhas.GetEdita(numLinha).Desconto1 = item.Desconto1;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PriEngine.Platform.MensagensDialogos.MostraErro(
                    "Aconteceu um erro durante um processo",
                    StdBE100.StdBETipos.IconId.PRI_Critico,
                    ex.Message,
                    ex
                );
            }
        }
    }
}
