#' @title Entrada de NF
#' @name fNFEntrada
#'
#' @description Uma (incrivel) funcao para fazer um xyplot
#'     personalizado. Usando o pacote lattice.
#'
#'
#' @details Utilize este campo para escrever detalhes mais tecnicos da
#'     sua funcao (se necessario), ou para detalhar melhor como
#'     utilizar determinados argumentos.
#'
#' @return 2 relatórios.
#'
#' @author Gelson Campos
#'
#'
#' @examples
#' fNFEntrada()
#'
#'
#' @export

fNFEntrada = function(){



  rm(list = ls())

  library(NFEntrada)

# LIBRARY
  library(odbc)
  library(httr)
  library(jsonlite)
  library(dplyr)
  library(RCurl)
  library(knitr)
  library(rmarkdown)
  library(markdown)
  library(kableExtra)
  library(ggthemes)
  library(ggplot2)
  library(gridExtra)
  library(scales)
  library(curl)
  library(XML)
  library(rvest)
  library(rjson)
  library(stringr)
  library(ggmap)
  library(googleway)
  library(bizdays)
  library(caret)
  library(h2o)
  library(caTools)
  library(GGally)
  #library(lubridate)
  library(leaflet)
  library(openxlsx)
  #library(mailR)
  library(blastula)
  library(keyring)
  library(glue)


# FUNCOES USADAS NO PROJETO --------------------------------------------------------------------------------------------
  #source(
  #  "J:/Suporte/DadosR/S_Funcoes.R",
  #  encoding = "UTF-8",
  # continue.echo = getOption("continue")
  #)

pasta_XML = "J:/XML/"
pasta_GLOG = "J:/Recebimento_Docs/"
pasta_SUPORTE = "J:/Suporte/"


message("CONECTANDO BANCO DE DADOS")
con = dbConnect(
  odbc::odbc(),
  Driver = "SQL Server",
  Server = "192.168.0.10",
  Database = "VSatNeotermica",
  UID = "POWERBI",
  PWD = "neotermica"
)
message("CONECTANDO BANCO DE DADOS - OK")

message("DOWNLOAD ARQUIVOS BASE")
base_NCM = read.xlsx(paste(pasta_SUPORTE,"NCM.xlsx",sep=""),
                     sheet = 1,
                     colNames = TRUE)


email_DIR = "luiz.marcomini@neotermica.com.br"
email_LOG = "barbara.chaves@neotermica.com.br"
email_COM = "compras@neotermica.com.br"
email_FIS = "recebimento@neotermica.com.br"

CD_FRETE = data.frame(
  "CD_FRETE" = c(0, 1, 2, 3, 4, 9),
  "Desc_FRETE" = c(
    "Contratação por conta do Remetente",
    "Contratação por conta do Destinatário",
    "Contratação por conta de Terceiros",
    "Transporte Próprio por conta do Remetente",
    "Transporte Próprio por conta do Destinatário",
    "Sem Ocorrência de Transporte"
  ),
  "Tipo_FRETE" = c("CIF",
                   "FOB",
                   "TER",
                   "PCR",
                   "PCD",
                   "SOT")
)


CST_IPI = data.frame(
  "CST_IPI" = c(
    "0",
    "1",
    "2",
    "3",
    "4",
    "5",
    "49",
    "50",
    "51",
    "52",
    "53",
    "54",
    "55",
    "99"
  ),
  "CST_IPI_Desc" = c(
    "Entrada com Recuperação de Crédito",
    "Entrada Tributável com Alíquota Zero",
    "Entrada Isenta",
    "Entrada Não-Tributada",
    "Entrada Imune",
    "Entrada com Suspensão",
    "Outras Entradas",
    "Saída Tributada",
    "Saída Tributável com Alíquota Zero",
    "Saída Isenta",
    "Saída Não-Tributada",
    "Saída Imune",
    "Saída com Suspensão",
    "Outras Saídas"
  )
)


CST_ICMS_SIT = data.frame(
  "CST_ICMS" = c("0", "1", "2", "3", "4", "5", "6", "7", "8"),
  "CST_ICMS_Desc" = c(
    "Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8",
    "Estrangeira - Importação direta, exceto a indicada no código 6",
    "Estrangeira - Adquirida no mercado interno, exceto a indicada no código 7",
    "Nacional, mercadoria ou bem com Conteúdo de Importação superior a 40% (quarenta por cento) e igual ou inferior a 70% (setenta por cento)",
    "Nacional, cuja produção tenha sido feita em conformidade com os processos produtivos básicos de que tratam o Decreto-Lei nÂº 288/1967 , e as Leis nÂºs 8.248/1991, 8.387/1991, 10.176/2001 e 11.484/2007",
    "Nacional, mercadoria ou bem com Conteúdo de Importação inferior ou igual a 40%",
    "Estrangeira - Importação direta, sem similar nacional, constante em lista de Resolução Camex e gás natural",
    "Estrangeira - Adquirida no mercado interno, sem similar nacional, constante em lista de Resolução Camex e gás natural",
    "Nacional - Mercadoria ou bem com Conteúdo de Importação superior a 70% (setenta por cento)"
  ),
  "CST_ICMS_Ali" = c(
    "7%/12%",
    "4%",
    "4%",
    "4%",
    "7%/12%",
    "7%/12%",
    "7%/12%",
    "7%/12%",
    "4%"
  )
)


CST_NEO = data.frame(
  "CST_ICMS" = c("0", "1", "2", "3", "4", "5", "6", "7", "8"),
  "Nacionalidade" = c("NAC", "IMP", "IMP", "NAC", "NAC", "NAC", "IMP", "IMP", "NAC"),
  "Proced" = c("NAC", "IMP", "IND", "N40", "NPP", "NA4", "IDC", "IIC", "N70"),
  "CST_ICMS_Ali" = c(
    "7%/12%",
    "4%",
    "4%",
    "4%",
    "7%/12%",
    "7%/12%",
    "7%/12%",
    "7%/12%",
    "4%"
  ),
  "CST_ICMS_Desc" = c(
    "Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8",
    "Estrangeira - Importação direta, exceto a indicada no código 6",
    "Estrangeira - Adquirida no mercado interno, exceto a indicada no código 7",
    "Nacional, mercadoria ou bem com Conteúdo de Importação superior a 40% (quarenta por cento) e igual ou inferior a 70% (setenta por cento)",
    "Nacional, cuja produção tenha sido feita em conformidade com os processos produtivos básicos de que tratam o Decreto-Lei nÂº 288/1967 , e as Leis nºs 8.248/1991, 8.387/1991, 10.176/2001 e 11.484/2007",
    "Nacional, mercadoria ou bem com Conteúdo de Importação inferior ou igual a 40%",
    "Estrangeira - Importação direta, sem similar nacional, constante em lista de Resolução Camex e gás natural",
    "Estrangeira - Adquirida no mercado interno, sem similar nacional, constante em lista de Resolução Camex e gás natural",
    "Nacional - Mercadoria ou bem com Conteúdo de Importação superior a 70% (setenta por cento)"
  )
)


CST_ICMS_TRI = data.frame(
  "CST_ICMS" = c("00", "10", "20", "30", "40", "41", "50", "51", "60", "70", "90"),
  "CST_ICMS_Desc" = c(
    "Tributada integralmente",
    "Tributada e com cobrança do ICMS por substituição tributária",
    "Com redução de base de cálculo",
    "Isenta ou nÃ£o tributada e com cobrança do ICMS por substituição tributária",
    "Isenta",
    "NÃ£o tributada",
    "SuspensÃ£o",
    "Diferimento",
    "ICMS cobrado anteriormente por substituição tributária",
    "Com redução de base de cálculo e cobrança do ICMS por substituição tributária",
    "Outras"
  )
)

CRT = data.frame(
  "CRT" = c("1", "2", "3"),
  "CRT_Desc" = c(
    "Simples Nacional",
    "Simples Nacional - Excesso de sublimite da receita bruta",
    "Regime Normal"
  )
)

message("DOWNLOAD ARQUIVOS BASE - OK")

consulta_lote = f_lote()

if (consulta_lote == "L") {
  lote_dados = read.csv2(
    paste(pasta_SUPORTE, "LOTE.csv", sep = ""),
    header = TRUE,
    sep = ";",
    dec = ",",
    fill = TRUE,
    na.strings = "N/A"
  )

  #write.table(lote_dados, file = paste(pasta_SUPORTE,"LOTE.csv",sep=""), sep = "|", na = "", quote = TRUE, row.names = FALSE)


  lote_dados$XML = str_replace_all(fcnpj(lote_dados$XML), " ", "")
  lote_dados$OC = str_replace_all(fcnpj(lote_dados$OC), " ", "")
  lote_dados$EMAIL = str_to_upper(str_replace_all(lote_dados$EMAIL, " ", ""))

} else {
  XML_NF = fxml_coletar()
  OC = fOC_coletar()
  envio_email = femail_coletar()

  lote_dados = data.frame("XML" = XML_NF,
                          "OC" = OC,
                          "EMAIL" = envio_email)

}

consulta_tamanho = nrow(lote_dados)

lote_dados_final = lote_dados
lote_dados_final$Prob_XML = ""
lote_dados_final$Prob_OC = ""

for (var_consulta in 1:consulta_tamanho) {



  tryCatch( {



    rm(list = setdiff(
      ls(),
      c(
        "lote_dados",
        "var_consulta",
        "consulta_tamanho",
        "pasta_SUPORTE",
        "base_NCM",
        "CD_FRETE",
        "CRT",
        "CST_IPI",
        "CST_ICMS_SIT",
        "CST_NEO",
        "CST_ICMS_TRI",
        "pasta_XML",
        "pasta_GLOG",
        "con",
        "lote_dados_final",
        "email_DIR",
        "email_LOG",
        "email_FIS"
      )
    ))

    email_DIR = "luiz.marcomini@neotermica.com.br"
    email_LOG = "barbara.chaves@neotermica.com.br"
    email_COM = "compras@neotermica.com.br"
    email_FIS = "recebimento@neotermica.com.br"

    nao_rodou = 0
    DANFE = "Arquivo não identificado!"

    #source(
    #  "J:/Suporte/DadosR/S_Funcoes.R",
    #  encoding = "UTF-8",
    #  continue.echo = getOption("continue")
    #)

    erro_no_xml = 0
    erro_na_OC = 0

    XML_NF = ifelse(lote_dados$XML[var_consulta] == 'NA' |
                      is.na(lote_dados$XML[var_consulta]) == TRUE,
                    0,
                    lote_dados$XML[var_consulta])
    OC = ifelse(lote_dados$OC[var_consulta] == 'NA' |
                  is.na(lote_dados$OC[var_consulta]) == TRUE,
                0,
                lote_dados$OC[var_consulta])
    envio_email =  ifelse(
      lote_dados$EMAIL[var_consulta] == 'NA' |
        is.na(lote_dados$EMAIL[var_consulta]) == TRUE,
      "N",
      lote_dados$EMAIL[var_consulta]
    )

    message(
      paste(
        "!!! REALIZANDO CONSULTA !!! CHAVE: ",
        XML_NF,
        " | OC: ",
        OC,
        " | EMAIL: ",
        envio_email,
        sep = ""
      )
    )


    if (str_count(XML_NF) != 44) {
      message(paste("!!! ERRO NO XML, TAMANHO INCORRETO !!!", XML_NF, sep = " - "))

      erro_no_xml = 1

    }

    lista_arquivos_xml = list.files(pasta_XML)

    XML_NF_esta_na_lista = paste(XML_NF, ".xml", sep = "") %in% lista_arquivos_xml

    DANFE_NF_esta_na_lista = paste(XML_NF, ".pdf", sep = "") %in% lista_arquivos_xml

    if (DANFE_NF_esta_na_lista == T){

      DANFE = paste("file:///",pasta_XML,XML_NF,".pdf",sep="")

    }


    if (XML_NF_esta_na_lista == FALSE) {
      message(paste("!!! ERRO NO XML, ARQUIVO NÃO ENCONTRADO !!!", XML_NF, sep =
                      " - "))

      erro_no_xml = 1

    }


    if (str_count(OC) != 4) {
      message(paste("!!! ERRO NA ORDEM DE COMPRA, TAMANHO INCORRETO !!!", OC, sep =
                      " - "))

      erro_na_OC = 1

    }




    if (envio_email != "S" & envio_email != "N") {
      envio_email = "N"

      message("!!! ERRO NA DEFINIÇÃO DO ENVIO DE E-MAIL. DESTACADO O PADRÃO N !!!")


    }

    lote_dados_final$Prob_XML[var_consulta] = ifelse(erro_no_xml == 1, "ERRO", "OK")
    lote_dados_final$Prob_OC[var_consulta] = ifelse(erro_na_OC == 1, "ERRO", "OK")


    if (erro_na_OC == 0 & erro_no_xml == 0) {
      message("BUSCANDO DADOS NO BANCO DE DADOS")




      SoliCompra = odbc::dbGetQuery(
        con,
        paste(
          "
  SELECT Materiais.cd_Referencia AS cd_Referencia
  , Materiais.ds_Prod AS ds_Prod
  , Materiais.cd_Prod AS cd_Prod
  , Solicita_Compras.Quantidade AS Quantidade
  , Solicitante.cdEnt AS cdSolicitante
  , Solicitante.Nome AS NomeSolicitante
  , Solicita_Compras.dt_Solicita AS dt_Solicita
  , Solicita_Compras.DataCompra AS DataCompra
  , Solicita_Compras.TipoSC AS TipoSC
  , Solicita_Compras.id_Sc AS idReg
  , Solicita_Compras.cd_Sc AS cd_Sc
  , Solicita_Compras.dt_InicioCompras AS dt_InicioCompras
  , Solicita_Compras.StatusSC AS StatusSC
  , Empresa.cdEnt AS cdEmpresa
  , Empresa.Nome AS NomeEmpresa
  , 1 AS ContadorSC
  , (Case when Solicita_Compras.dtPrevEntrega - isnull( Solicita_Compras.DataCompra , getdate()) >=0 then 1 else 0 end) AS QtdSolicAtendOK
  , (Case when Solicita_Compras.dtPrevEntrega - isnull( Solicita_Compras.DataCompra , getdate()) <0 then 1 else 0 end) AS QtdSolicAtraso
  , (Case when Solicita_Compras.DataCompra is null then 1 else 0 end) AS QtdSCaberto
  , (Case when DateDiff(DD, Solicita_Compras.dt_Solicita,Solicita_Compras.dtPrevEntrega) < 3 then 1 else 0 end) AS QtdScUrgente
  , (case when Solicita_Compras.StatusSC = 'LIQ' then 0 else (case when datediff(Dd,isnull(Solicita_Compras.DataCompra,getdate()),Solicita_Compras.dtPrevEntrega)<0 then datediff(DD,isnull( Solicita_Compras.DataCompra,getdate()),Solicita_Compras.dtPrevEntrega) else 0 end) end) AS QtdDiasAtraso
  , Comprador.cdEnt AS cdComprador
  , Comprador.Nome AS NomeComprador
  , Comprador.Apelido AS ApelidoComprador
  , Ordem_Compras.cdOC AS cdOC
  , Mapa_Ret_Precos.cd_MapaCompras AS cd_MapaCompras
  , Unid_Medidas.cd_UnidMed AS cd_UnidMed
  , ISNULL(SCOwner.Cod, 0) AS CodOwner
  , Solicita_Compras.Observacao AS Observacao

   FROM Solicita_Compras AS Solicita_Compras  WITH(NOLOCK)
  INNER JOIN HistFinanceiro AS HistFinanceiro  WITH(NOLOCK) ON (HistFinanceiro.IdHist = Solicita_Compras.IdHist)
  INNER JOIN Materiais AS Materiais  WITH(NOLOCK) ON (Materiais.id_Produto = Solicita_Compras.id_Produto)
  INNER JOIN PlanoCusteio AS PlanoCusteio  WITH(NOLOCK) ON (PlanoCusteio.id_CCusto = Solicita_Compras.CCustoAlocacao)
  INNER JOIN Entidade AS Empresa  WITH(NOLOCK) ON (Empresa.Id_Ent = Solicita_Compras.id_empr)
  INNER JOIN Entidade AS Solicitante  WITH(NOLOCK) ON (Solicitante.Id_Ent = Solicita_Compras.id_Prof_Intern)
  INNER JOIN Prof_Internos AS Prof_Internos  WITH(NOLOCK) ON (Prof_Internos.Id_Ent = Solicita_Compras.id_Prof_Intern)
  LEFT JOIN Ordem_Compras AS Ordem_Compras  WITH(NOLOCK) ON (Ordem_Compras.id_OC = Solicita_Compras.id_OC AND Ordem_Compras.PierSitReg = 'ATV')
  LEFT JOIN Mapa_Ret_Precos AS Mapa_Ret_Precos  WITH(NOLOCK) ON (Mapa_Ret_Precos.id_MapaCompras = Solicita_Compras.id_MapaCompras AND Mapa_Ret_Precos.PierSitReg = 'ATV')
  INNER JOIN SubGrupoProduto AS SubGrupoProduto  WITH(NOLOCK) ON (SubGrupoProduto.id_SubGrupoPrd = Materiais.id_SubGrupoPrd)
  INNER JOIN GrupoProduto AS GrupoProduto  WITH(NOLOCK) ON (GrupoProduto.id_grpProd = SubGrupoProduto.id_grpProd)
  INNER JOIN CatGrupo AS CatGrupo  WITH(NOLOCK) ON (CatGrupo.id_catGrupo = GrupoProduto.id_catGrupo)
  INNER JOIN ParamEmpresaMateriais AS ParamEmpresaMateriais  WITH(NOLOCK) ON (ParamEmpresaMateriais.Id_Ent = Empresa.Id_Ent AND ParamEmpresaMateriais.id_Produto = Materiais.id_Produto AND ParamEmpresaMateriais.PierSitReg = 'ATV')
  LEFT JOIN Entidade AS Comprador  WITH(NOLOCK) ON (Comprador.Id_Ent = Mapa_Ret_Precos.id_Comprador)
  INNER JOIN Unid_Medidas AS Unid_Medidas  WITH(NOLOCK) ON (Unid_Medidas.id_unidMed = Solicita_Compras.id_unidMed AND Unid_Medidas.PierSitReg = 'ATV')
  OUTER APPLY (SELECT SC.cd_Sc AS Cod
    FROM Solicita_Compras AS SC WITH(NOLOCK)
    WHERE SC.id_Sc = Solicita_Compras.id_RegOwner AND SC.PierSitReg = 'ATV') AS SCOwner
  WHERE (Solicita_Compras.PierSitReg = 'ATV') AND Ordem_Compras.cdOC = '",
          OC,
          "'

  ORDER BY cdEmpresa
  , dt_Solicita
  , cdSolicitante
  , cd_Sc",
          sep = ""
        )
      )


      OrdemCompra = odbc::dbGetQuery(
        con,
        paste(
          "
  SELECT Ordem_Compras.cdOC AS cdOC
  , Ordem_Compras.DataCompra AS DataCompra
  , Ordem_Compras.TipoTransacaoCompra AS TipoTransacaoCompra
  , Ordem_Compras.StatusOC AS StatusOC
  , Comprador.cdEnt AS CodComprador
  , Comprador.Nome AS NomeComprador
  , Fornecedor.cdEnt AS CodFornecedor
  , Fornecedor.Nome AS NomeFornecedor
  , Fornecedor_cnpj.CNPJ AS CNPJ
  , Cond_Pagto.cd_cdPagt AS cd_cdPagt
  , Cond_Pagto.ds_cdPagt AS ds_cdPagt
  , Ordem_Compras.Cif_Fob AS Cif_Fob
  , Materiais.cd_Referencia AS cd_Referencia
  , Materiais.ds_Prod AS ds_Prod
  , Det_Ord_Compra.qtdItem AS qtdItem
  , Det_Ord_Compra.Qtd_Recebida AS Qtd_Recebida
  , Det_Ord_Compra.QtdCancelado AS QtdCancelado
  , Det_Ord_Compra.QtdSaldo AS QtdSaldo
  , Det_Ord_Compra.PrecoUnitario AS PrecoUnitario
  , Det_Ord_Compra.VlTotal AS VlTotal
  , Det_Ord_Compra.ValorFinal AS ValorFinalItem
  , Det_Ord_Compra.Ordenacao AS Ordenacao
  , Det_Ord_Compra.vlrIPI AS vlrIPI
  , Det_Ord_Compra.VlrIcms AS VlrIcms
  , Det_Ord_Compra.VlrBaseSubst AS VlrBaseSubst
  , Det_Ord_Compra.vlrICMSSubst AS vlrICMSSubst
  , Det_Ord_Compra.ValorFCP AS ValorFCP
  , Det_Ord_Compra.ValorAdicional AS ValorAdicional
  , Naturezas_Operacao.ds_Natop AS ds_Natop
  , Naturezas_Operacao.cd_Natop AS cd_Natop
  , RelacFornProduto.cdProdForn AS cdProdForn
  , RelacFornProduto.DsProdForn AS DsProdForn
  , RelacFornProduto.Sta_OrigemMat AS CST_Origem_Forn
  , Det_Ord_Compra.Ali_Icms AS Ali_Icms
  , Det_Ord_Compra.Ali_IPI AS Ali_IPI
  , Naturezas_Operacao.Cod_CFOP AS Cod_CFOP
  , Naturezas_Operacao.Sta_UsoConsumo AS Sta_UsoConsumo
  , Naturezas_Operacao.TipodeOperacao AS TipodeOperacao
  , Naturezas_Operacao.TabTributacaoICMS AS TabTributacaoICMS
  , Ordem_Compras.data_liberacao AS data_liberacao
  , Det_Ord_Compra.NumeroItemPedidoCliente AS NumeroItemPedidoCliente
  , Materiais.Sta_ImportNacion AS Sta_ImportNacion
  , Materiais.Sta_TipoProd AS Sta_TipoProd
  , Naturezas_Operacao.ClassICMS AS ClassICMS
  , Cond_Pagto.PrazoMedio AS PrazoMedio
  , Det_Ord_Compra.Saldo AS Saldo
  , Det_Ord_Compra.Vlr_Recebido AS Vlr_Recebido
  , (CASE WHEN Det_Ord_Compra.VlrIcms = 0 THEN 0 ELSE(Det_Ord_Compra.Ali_Icms  * Det_Ord_Compra.QtdSaldo * Det_Ord_Compra.PrecoUnitario)/100 END)
   AS VlrIcms_Saldo
  , (CASE WHEN Det_Ord_Compra.vlrIPI = 0 THEN 0 ELSE(Det_Ord_Compra.Ali_IPI  * Det_Ord_Compra.QtdSaldo * Det_Ord_Compra.PrecoUnitario)/100 END) AS VlrIPISaldo
  , (CASE WHEN Det_Ord_Compra.VlrIcms = 0 THEN 0 ELSE(Det_Ord_Compra.Ali_Icms * Det_Ord_Compra.Qtd_Recebida * Det_Ord_Compra.PrecoUnitario)/100 END)
   AS VlrIcmsRecebido
  , (CASE WHEN Det_Ord_Compra.vlrIPI = 0 THEN 0 ELSE(Det_Ord_Compra.Ali_IPI * Det_Ord_Compra.Qtd_Recebida * Det_Ord_Compra.PrecoUnitario)/100 END) AS VlrIPIRecebido
  , (CASE WHEN Det_Ord_Compra.vlrIPI = 0 THEN 0 ELSE(Det_Ord_Compra.Ali_IPI  * Det_Ord_Compra.QtdSaldo * Det_Ord_Compra.PrecoUnitario)/100 END) + (Det_Ord_Compra.QtdSaldo * Det_Ord_Compra.PrecoUnitario)
   AS VlrTotalSaldo
  , (CASE WHEN Det_Ord_Compra.vlrIPI = 0 THEN 0 ELSE(Det_Ord_Compra.Ali_IPI  * Det_Ord_Compra.Qtd_Recebida * Det_Ord_Compra.PrecoUnitario)/100 END) + (Det_Ord_Compra.Qtd_Recebida * Det_Ord_Compra.PrecoUnitario)
   AS VlrRecebidoTotal
  , Naturezas_Operacao.ProcedenciaMat AS ProcedenciaMat
  , ParamEmpresaMateriais.Sta_OrigemMat AS Sta_OrigemMat_PE
  , Unid_Medidas.cd_UnidMed AS cd_UnidMed
  , HistFinanceiro.cd_HistFin AS cd_HistFin
  , HistFinanceiro.dshistfin AS dshistfin
  , HistFinanceiro.Tipo_Hist_Fin AS Tipo_Hist_Fin
  , NBM.ds_NBM AS ds_NBM
  , NBM.NCM AS NCM
  , NBM.CodNBM AS CodNBM
  , Det_Ord_Compra.PrecoUnitario * Det_Ord_Compra.QtdSaldo AS ProdTotalVlr
  , Det_Ord_Compra.PrecoUnitario * Det_Ord_Compra.Qtd_Recebida AS ProdValorRec
  , (CASE WHEN Det_Ord_Compra.QtdSaldo = 0 THEN 'TOTAL RECEBIDO' ELSE '' END) AS NeoInformacao
  , MVA.Ali_Icms AS MVA_ALI_ICMS
  , MVA.Aliq_MVA AS MVA_Aliq_MVA
  , Fornecedor_UF.cd_SglEstado AS UF
  , FaixaICMS.fx_ICMS AS fx_ICMS_Saida
  , FaixaICMS.Ali_ICMS AS Ali_ICMS_Saida
  , CContabil.cd_cdreduzida AS CContabil
  , CContabil.ds_conta AS DSContabil
  , CConsumo.cd_cdreduzida AS CConsumo
  , CConsumo.ds_conta AS DSConsumo
  , Ordem_Compras.Observacao AS Observacao

   FROM Ordem_Compras AS Ordem_Compras  WITH(NOLOCK)
  INNER JOIN Det_Ord_Compra AS Det_Ord_Compra  WITH(NOLOCK) ON (Det_Ord_Compra.id_OC = Ordem_Compras.id_OC and Det_Ord_Compra.PierSitReg = 'ATV' AND Det_Ord_Compra.QtdSaldo > 0)
  INNER JOIN Entidade AS Empresa  WITH(NOLOCK) ON (Empresa.Id_Ent = Ordem_Compras.id_empr)
  INNER JOIN Entidade AS Comprador  WITH(NOLOCK) ON (Comprador.Id_Ent = Ordem_Compras.id_Comprador)
  INNER JOIN Entidade AS Fornecedor  WITH(NOLOCK) ON (Fornecedor.Id_Ent = Ordem_Compras.id_Forn)
  LEFT JOIN Entid_PessoasJur AS Fornecedor_cnpj  WITH(NOLOCK) ON (Fornecedor.Id_Ent = Fornecedor_cnpj.Id_Ent AND Fornecedor_cnpj.PierSitReg = 'ATV')
  LEFT JOIN Unid_Federacao AS Fornecedor_UF  WITH(NOLOCK) ON (Fornecedor_UF.id_SglEstado = Fornecedor_cnpj.id_SglEstado)
  INNER JOIN Cond_Pagto AS Cond_Pagto  WITH(NOLOCK) ON (Cond_Pagto.id_cdPagt = Ordem_Compras.id_cdPagt)
  INNER JOIN ContaFinanceira AS ContaFinanceira  WITH(NOLOCK) ON (ContaFinanceira.id_Cta_Financeira = Ordem_Compras.id_Cta_Financeira)
  INNER JOIN Materiais AS Materiais  WITH(NOLOCK) ON (Materiais.id_Produto = Det_Ord_Compra.id_Produto)
  INNER JOIN Naturezas_Operacao AS Naturezas_Operacao  WITH(NOLOCK) ON (Naturezas_Operacao.id_Natop = Det_Ord_Compra.id_Natop)
  LEFT JOIN RelacFornProduto AS RelacFornProduto  WITH(NOLOCK) ON (RelacFornProduto.id_Produto = Det_Ord_Compra.id_Produto and RelacFornProduto.Id_Ent = Ordem_Compras.id_Forn and RelacFornProduto.PierSitReg = 'ATV')
  LEFT JOIN ParamEmpresaMateriais AS ParamEmpresaMateriais  WITH(NOLOCK) ON (ParamEmpresaMateriais.id_Produto = Det_Ord_Compra.id_Produto and ParamEmpresaMateriais.Id_Ent = '1' and ParamEmpresaMateriais.PierSitReg = 'ATV')
  LEFT JOIN Unid_Medidas AS Unid_Medidas  WITH(NOLOCK) ON (Unid_Medidas.id_unidMed = Det_Ord_Compra.id_unidMed)
  LEFT JOIN Aliq_ICMS_Estados AS FaixaICMS  WITH(NOLOCK) ON (FaixaICMS.fx_ICMS = Materiais.fx_ICMS  AND FaixaICMS.id_SglEstado = Fornecedor_UF.id_SglEstado AND FaixaICMS.Entra_Sai = 'SAI' AND FaixaICMS.id_Empresa = '1')
  INNER JOIN HistFinanceiro AS HistFinanceiro  WITH(NOLOCK) ON (HistFinanceiro.IdHist = Ordem_Compras.IdHist)
  LEFT JOIN NBM AS NBM  WITH(NOLOCK) ON (NBM.id_NBM = Materiais.id_NBM)
  LEFT JOIN SubstTributMVA AS MVA WITH(NOLOCK) ON (MVA.id_NBM = NBM.id_NBM AND MVA.PierSitReg = 'ATV' AND MVA.id_SglEstado = '27' and MVA.EstadoCliente = 'S')
  LEFT JOIN PlanoCtacContab AS CContabil WITH(NOLOCK) ON (CContabil.id_ctcInterno = Materiais.id_ctcInterno and CContabil.PierSitReg = 'ATV')
  LEFT JOIN PlanoCtacContab AS CConsumo WITH(NOLOCK) ON (CConsumo.id_ctcInterno = Materiais.Id_CtConsumo and CConsumo.PierSitReg = 'ATV')
  WHERE (Ordem_Compras.PierSitReg = 'ATV') AND Ordem_Compras.cdOC = '",
          OC,
          "'

  ORDER BY Empresa.cdEnt
  , Ordem_Compras.cdOC",
          sep = ""
        )
      )


      if (nrow(OrdemCompra) == 0) {
        message(
          paste(
            "!!! ERRO NA ORDEM DE COMPRA, NÃO IDENTIFICADA OU NÃO EXISTE SALDO !!!",
            OC,
            sep = " - "
          )
        )

        erro_na_OC = 1

      } else {
        OC_CNPJ = str_trim(gsub('[[:punct:]]', '', max(OrdemCompra$CNPJ)))
        CNPJ_EX_XML = str_sub(XML_NF, 7, 20)
        if (OC_CNPJ != CNPJ_EX_XML) {
          message(
            paste(
              "!!! ERRO NA ORDEM DE COMPRA x XML. CNPJ DISTINTO NOS 2 DOCUMENTOS !!!",
              sep = " - "
            )
          )
          erro_na_OC = 1
        }

      }

      lote_dados_final$Prob_XML[var_consulta] = ifelse(erro_no_xml == 1, "ERRO", "OK")
      lote_dados_final$Prob_OC[var_consulta] = ifelse(erro_na_OC == 1, "ERRO", "OK")

      if (erro_na_OC == 0 & erro_no_xml == 0) {
        OC_CNPJ = str_trim(gsub('[[:punct:]]', '', max(OrdemCompra$CNPJ)))

        OC_razao_Social = str_trim(max(OrdemCompra$NomeFornecedor))

        cd_Cond_Pagto = str_trim(max(OrdemCompra$cd_cdPagt))

        OC_Status = str_trim(max(OrdemCompra$StatusOC))

        bd_Cond_Pagto = dbGetQuery(
          con,
          paste(
            "
    SELECT cd_cdPagt,
    ds_cdPagt,
    PrazoMedio,
    qtdDias
    FROM Cond_Pagto
    INNER JOIN DetalheCondPagto WITH(NOLOCK) ON (DetalheCondPagto.id_cdPagt = Cond_Pagto.id_cdPagt)
    WHERE Cond_Pagto.PierSitReg = 'ATV' AND DetalheCondPagto.PierSitReg = 'ATV' AND Cond_Pagto.cd_cdPagt = '",
            cd_Cond_Pagto,
            "'

    ORDER BY Cond_Pagto.cd_cdPagt",
            sep = ""
          )
        )


        bd_UM_Conv = dbGetQuery(
          con,
          paste(
            "
    SELECT Conv_UnidMed.unid_medOrig AS ID_UM_EMPRESA,
    Conv_UnidMed.unid_medOrig AS ID_UM_FORN,
    UM_EMPRESA.cd_UnidMed AS UM_EMPRESA,
    UM_FORN.cd_UnidMed AS UM_FORN,
    Fator_Conv,
    OperMatConvUnidMed
    FROM Conv_UnidMed
    LEFT JOIN Unid_Medidas AS UM_EMPRESA WITH(NOLOCK) ON (UM_EMPRESA.id_unidMed = Conv_UnidMed.unid_medOrig)
    LEFT JOIN Unid_Medidas AS UM_FORN WITH(NOLOCK) ON (UM_FORN.id_unidMed = Conv_UnidMed.unid_medDest)
    WHERE Conv_UnidMed.PierSitReg = 'ATV'",
            sep = ""
          )
        )

        bd_UM_Conv$chave_UM = paste(str_trim(bd_UM_Conv$UM_EMPRESA),
                                    str_trim(bd_UM_Conv$UM_FORN),
                                    sep = "-")

        bdCFOP = read.xlsx(
          "J:/Suporte/DadosR/CFOP.xlsx",
          sheet = 1,
          colNames = TRUE
        )
        bdCFOP = bdCFOP[, c("CFOP", "Descr")]
        colnames(bdCFOP) = c("CFOP", "Descr_CFOP")


        message("BUSCANDO XML")

        xml_salvar = xml2::read_xml(paste(pasta_XML, XML_NF, ".xml", sep = ""))

        ts = xmlParse(paste(pasta_XML, XML_NF, ".xml", sep = ""), encoding = "UTF-8")
        xml_lista <- xmlToList(ts)




        message("REALIZANDO ANÁLISE")


        ### DADOS GERAIS ----------------------------------------------------------

        cd_Forn_Areco = max(OrdemCompra$CodFornecedor)
        Comprador = str_trim(max(OrdemCompra$NomeComprador))
        Cond_Pagto = str_trim(max(OrdemCompra$ds_cdPagt))
        Desc_HistFin = str_trim(max(OrdemCompra$dshistfin))
        cd_HistFin = str_trim(max(OrdemCompra$cd_HistFin))
        Tipo_Tran = str_trim(max(OrdemCompra$TipoTransacaoCompra))
        Qtde_Item_OC = nrow(OrdemCompra)
        Data_Compra = as.Date(str_trim(max(OrdemCompra$DataCompra)), "%Y-%m-%d")

        frete_OC = str_trim(max(OrdemCompra$Cif_Fob))
        frete_OC = CD_FRETE[CD_FRETE$Tipo_FRETE == frete_OC, ]

        frete_OC = paste(frete_OC$CD_FRETE,
                         "|",
                         frete_OC$Desc_FRETE,
                         "|",
                         frete_OC$Tipo_FRETE,
                         sep = "")

        ObservacaoOC = max(OrdemCompra$Observacao)



        PrazoMedioOC = max(bd_Cond_Pagto$PrazoMedio)
        QtdeParcedlasOC = nrow(bd_Cond_Pagto)

        DadosNF = xml_lista$NFe$infNFe$ide

        DG_NF = DadosNF$nNF
        DG_NATOP = DadosNF$natOp
        DG_EMISSAO = str_sub(DadosNF$dhEmi, 1, 10)
        DG_USO = xml_lista$protNFe$infProt$xMotivo
        DG_CHAVE = xml_lista$protNFe$infProt$chNFe
        DG_CRT = xml_lista$NFe$infNFe$emit$CRT
        DG_CRT = subset(CRT,
                        CRT == DG_CRT)

        DG_infAdic = xml_lista$NFe$infNFe$infAdic

        TempoCompraNF = paste(as.numeric(as.Date(DG_EMISSAO, "%Y-%m-%d") - Data_Compra), "dias", sep =
                                " ")

        xml2::write_xml(xml_salvar,paste("Q:/NFE ENTRADA/",str_sub(DG_EMISSAO,1,4),"/",str_sub(DG_EMISSAO,1,4),"_",str_sub(DG_EMISSAO,6,7),"/INS_CONS/",XML_NF,".xml",sep=""),option = "as_xml",encoding = "UTF-8")


        ### EMITENTE ----------------------------------------------------------

        Emitente = xml_lista$NFe$infNFe$emit



        Forn_CNPJ = str_trim(gsub('[[:punct:]]', '', Emitente$CNPJ))
        Forn_RazaoSocial = str_to_upper(Emitente$xNome)
        Forn_IE = Emitente$IE
        Forn_IM = Emitente$IM
        Forn_CRT = Emitente$CRT
        Forn_UF = Emitente$enderEmit$UF

        ### DESTINATÁRIO ----------------------------------------------------------

        Destinatario = xml_lista$NFe$infNFe$dest


        Dest_CNPJ = Destinatario$CNPJ
        Dest_RazaoSocial = str_to_upper(Destinatario$xNome)
        Dest_IE = Destinatario$IE
        Dest_UF = Destinatario$enderDest$UF

        ### TRANSPORTADORA ----------------------------------------------------------

        Transportadora = xml_lista$NFe$infNFe$transp


        Trans_CNPJ = Transportadora$transporta$CNPJ
        Trans_RazaoSocial = str_to_upper(Transportadora$transporta$xNome)
        Trans_IE = Transportadora$transporta$IE
        Trans_UF = Transportadora$transporta$UF
        Trans_Placa = Transportadora$veicTransp$placa

        Trans_Frete = Transportadora$modFrete

        Trans_Frete = subset(CD_FRETE,
                             CD_FRETE == Trans_Frete)

        Trans_Frete_cd = Trans_Frete$CD_FRETE

        Trans_Frete = paste(
          Trans_Frete$CD_FRETE,
          "|",
          Trans_Frete$Desc_FRETE,
          "|",
          Trans_Frete$Tipo_FRETE,
          sep = ""
        )

        if (is.null(Transportadora$vol[1]$qVol)  == FALSE) {
          for (i in 1:100) {
            if (is.null(Transportadora[i]$vol$qVol) == FALSE) {
              fim_frete = i

            }
          }

          for (i in 100:1) {
            if (is.null(Transportadora[i]$vol$qVol) == FALSE) {
              inic_frete = i

            }
          }


          Frete_NF_Quadro = data.frame()

          for (u in inic_frete:fim_frete) {
            Frete_Inter = data.frame(
              "qVol" = Transportadora[u]$vol$qVol,
              "pesoL" = Transportadora[u]$vol$pesoL,
              "pesoB" = Transportadora[u]$vol$pesoB
            )

            Frete_NF_Quadro = rbind(Frete_NF_Quadro, Frete_Inter)

            Frete_Inter = NULL

          }

          Trans_Volume = sum(as.numeric(Frete_NF_Quadro$qVol))

        } else {
          Frete_NF_Quadro = data.frame()

          Trans_Volume = 0

        }


        ### PAGTO ----------------------------------------------------------

        Pagto = xml_lista$NFe$infNFe$cobr

        inic_pag = 0

        fim_pag = 0

        for (i in 1:20) {
          if (is.null(Pagto[i]$dup$nDup) == FALSE) {
            fim_pag = i

          }
        }

        for (i in 20:1) {
          if (is.null(Pagto[i]$dup$nDup) == FALSE) {
            inic_pag = i

          }
        }

        Pagto_NF_Quadro = data.frame()

        Pagto_nao_dest = 0

        if (inic_pag == 0 & fim_pag == 0) {
          Pagto_NF_Quadro = data.frame(
            "Parcela" = "1",
            "Vencimento" = "1900-01-01",
            "Valor" = Pagto$fat$vLiq
          )
          Pagto_nao_dest = 1

        } else {
          for (u in inic_pag:fim_pag) {
            PagtoNF_Inter = data.frame(
              "Parcela" = Pagto[u]$dup$nDup,
              "Vencimento" = Pagto[u]$dup$dVenc,
              "Valor" = Pagto[u]$dup$vDup
            )

            Pagto_NF_Quadro = rbind(Pagto_NF_Quadro, PagtoNF_Inter)

            PagtoNF_Inter = NULL

          }
        }




        Qtde_ParcelasNF = nrow(Pagto_NF_Quadro)

        VlrFIN_NF = sum(as.numeric(Pagto_NF_Quadro$Valor))

        VlrFIN_Desc_NF = Pagto$fat$vDesc

        Pagto_NF_Quadro$DiasEmissao =  as.numeric(
          as.Date(Pagto_NF_Quadro$Vencimento, "%Y-%m-%d")  - as.Date(DG_EMISSAO, "%Y-%m-%d")
        )

        PrazoMedioNF = sum(Pagto_NF_Quadro$DiasEmissao) / Qtde_ParcelasNF



        ### INFORMAÇÕES ADICIONAIS ----------------------------------------------------------

        InfAdic = xml_lista$NFe$infNFe$infAdic


        ### VALORES TOTAIS ----------------------------------------------------------

        Total = xml_lista$NFe$infNFe$total

        vICMS_Total = Total$ICMSTot$vICMS
        vST_Total = Total$ICMSTot$vST
        vProd_Total = Total$ICMSTot$vProd
        vFrete_Total = Total$ICMSTot$vFrete
        vDesc_Total = Total$ICMSTot$vDesc
        vIPI_Total = Total$ICMSTot$vIPI
        vNF_Total = Total$ICMSTot$vNF

        ### PRODUTOS ----------------------------------------------------------

        FINAL = ""
        PRODUTOS_NF = data.frame("ID" = "")
        ID = 0

        for (i in 1:100) {
          if (is.null(xml_lista$NFe$infNFe[i]$det) == FALSE) {
            fim = i

          }
        }

        for (i in 100:1) {
          if (is.null(xml_lista$NFe$infNFe[i]$det) == FALSE) {
            inic = i

          }
        }


        for (u in inic:fim) {
          PROD = xml_lista$NFe$infNFe[u]$det$prod

          PROD = PROD[names(PROD) != "rastro"]

          PROD = as.data.frame(do.call(cbind, PROD))

          PROD = PROD[, colnames(PROD) != "rastro"]

          row.names(PROD) = NULL

          ICMS = as.data.frame(do.call(rbind, xml_lista$NFe$infNFe[u]$det$imposto$ICMS))
          for (i in 1:ncol(ICMS)) {
            colnames(ICMS)[i] = paste("ICMS_", colnames(ICMS)[i], sep = "")
          }
          if (is.null(xml_lista$NFe$infNFe[u]$det$imposto$IPI$IPITrib) == TRUE) {
            IPI = as.data.frame(do.call(
              cbind,
              xml_lista$NFe$infNFe[u]$det$imposto$IPI$IPINT
            ))

          } else {
            IPI = as.data.frame(do.call(
              cbind,
              xml_lista$NFe$infNFe[u]$det$imposto$IPI$IPITrib
            ))

          }

          for (i in 1:ncol(IPI)) {
            colnames(IPI)[i] = paste("IPI_", colnames(IPI)[i], sep = "")
          }

          PIS = as.data.frame(do.call(rbind, xml_lista$NFe$infNFe[u]$det$imposto$PIS))
          for (i in 1:ncol(PIS)) {
            colnames(PIS)[i] = paste("PIS_", colnames(PIS)[i], sep = "")
          }

          COFINS = as.data.frame(do.call(rbind, xml_lista$NFe$infNFe[u]$det$imposto$COFINS))
          for (i in 1:ncol(COFINS)) {
            colnames(COFINS)[i] = paste("COFINS_", colnames(COFINS)[i], sep = "")
          }

          FINAL = cbind(PROD, ICMS, IPI, PIS, COFINS)
          ID = ID + 1
          FINAL$ID = ID

          PRODUTOS_NF = merge(PRODUTOS_NF, FINAL, all = T)

          FINAL = ""

        }


        PRODUTOS_NF = PRODUTOS_NF[PRODUTOS_NF$ID > 0, ]

        PRODUTOS_NF = as.data.frame(PRODUTOS_NF)


        PRODUTOS_MODELO = data.frame(
          "ID" = 1,
          "cProd" = 1,
          "cEAN" = 1,
          "xProd" = 1,
          "NCM" = 1,
          "CFOP" = 1,
          "uCom" = 1,
          "qCom" = 1,
          "vUnCom" = 1,
          "vProd" = 1,
          "cEANTrib" = 1,
          "uTrib" = 1,
          "qTrib" = 1,
          "vUnTrib" = 1,
          "indTot" = 1,
          "ICMS_orig" = 1,
          "ICMS_CST" = 1,
          "ICMS_modBC" = 1,
          "ICMS_vBC" = 1,
          "ICMS_pICMS" = 1,
          "ICMS_vICMS" = 1,
          "IPI_CST" = 1,
          "IPI_pIPI" = 1,
          "IPI_vIPI" = 1,
          "PIS_CST" = 1,
          "PIS_vBC" = 1,
          "PIS_pPIS" = 1,
          "PIS_vPIS" = 1,
          "COFINS_CST" = 1,
          "COFINS_vBC" = 1,
          "COFINS_pCOFINS" = 1,
          "COFINS_vCOFINS" = 1,
          "ICMS_modBCST" = 1,
          "ICMS_vBCST" = 1,
          "ICMS_pICMSST" = 1,
          "ICMS_vICMSST" = 1,
          "vDesc" = 1
        )


        PRODUTOS_MODELO = PRODUTOS_MODELO[-1, ]

        PRODUTOS_NF = merge(PRODUTOS_NF, PRODUTOS_MODELO, all.x = TRUE)

        ########################################### TESTES -------------------------------------- #############################################


        ########################################### CÓDIGO DE PRODUTO USADO PARA MAIS DE ITEM NA NF

        base_nf = data.frame(
          "Ref_Forn" = str_trim(PRODUTOS_NF$cProd),
          "Desc_Forn" = str_trim(PRODUTOS_NF$xProd),
          "UM_Prod_Forn" = str_trim(PRODUTOS_NF$uCom),
          "NF_QtdeProd" = as.numeric(PRODUTOS_NF$qCom)
        )

        base_nf = base_nf %>%
          group_by(Ref_Forn) %>%
          summarise("Qtde de Prod Cód" = length(Ref_Forn))

        base_nf = as.data.frame(base_nf)

        ref_qtde_acima1 = base_nf[base_nf$`Qtde de Prod Cód` > 1, ]

        ########################################### QTDE

        base_dest = data.frame(
          "Ref_Prod" = str_trim(OrdemCompra$cd_Referencia),
          "Desc_Prod" = str_trim(OrdemCompra$ds_Prod),
          "UM_Prod" = str_trim(OrdemCompra$cd_UnidMed),
          "Ref_Forn" = str_trim(OrdemCompra$cdProdForn),
          "OC_QtdeProd" = OrdemCompra$QtdSaldo
        )

        base_dest = base_dest %>%
          group_by(Ref_Prod, Desc_Prod, UM_Prod, Ref_Forn) %>%
          summarise("OC_QtdeProd" = sum(OC_QtdeProd))


        base_nf = data.frame(
          "Ref_Forn" = str_trim(PRODUTOS_NF$cProd),
          "Desc_Prod_Forn" = str_trim(PRODUTOS_NF$xProd),
          "UM_Prod_Forn" = str_trim(PRODUTOS_NF$uCom),
          "NF_QtdeProd" = as.numeric(PRODUTOS_NF$qCom)
        )

        base_nf = base_nf %>%
          group_by(Ref_Forn, Desc_Prod_Forn, UM_Prod_Forn) %>%
          summarise("NF_QtdeProd" = sum(NF_QtdeProd))



        nao_relacionados = merge(
          x = base_nf,
          y = base_dest,
          by.x = "Ref_Forn",
          by.y = "Ref_Forn",
          all.x = TRUE
        )
        nao_relacionados = nao_relacionados[is.na(nao_relacionados$OC_QtdeProd) ==
                                              TRUE, ]
        qtde_nao_relacionados = nrow(nao_relacionados)

        nao_entregues = merge(
          x = base_nf,
          y = base_dest,
          by.x = "Ref_Forn",
          by.y = "Ref_Forn",
          all.y = TRUE
        )
        nao_entregues = nao_entregues[is.na(nao_entregues$NF_QtdeProd) == TRUE, ]
        qtde_nao_nao_entregues = nrow(nao_entregues)

        base_teste_qtde_oc = merge(base_nf, base_dest, by.x = "Ref_Forn", by.y = "Ref_Forn")


        base_teste_qtde_oc$Dif_qtde = base_teste_qtde_oc$NF_QtdeProd - base_teste_qtde_oc$OC_QtdeProd

        tem_SC = merge(SoliCompra,
                       base_teste_qtde_oc,
                       by.x = "cd_Referencia",
                       by.y = "Ref_Prod")



        ########################################### NCM

        base_dest = data.frame(
          "Ref_Prod" = str_trim(OrdemCompra$cd_Referencia),
          "Desc_Prod" = str_trim(OrdemCompra$ds_Prod),
          "NCM" = str_trim(gsub('[[:punct:]]', '', OrdemCompra$NCM)),
          "Ref_Forn" = str_trim(OrdemCompra$cdProdForn)
        )

        base_dest = base_dest %>%
          group_by(Ref_Prod, Desc_Prod, NCM, Ref_Forn)


        base_nf = data.frame("Ref_Forn" = str_trim(PRODUTOS_NF$cProd),
                             "Forn_NCM" = str_trim(gsub('[[:punct:]]', '', PRODUTOS_NF$NCM)))

        base_nf = base_nf %>%
          group_by(Ref_Forn, Forn_NCM)

        NCM = merge(base_nf, base_dest, by.x = "Ref_Forn", by.y = "Ref_Forn")

        NCM = NCM %>%
          group_by(Ref_Prod, Desc_Prod, Ref_Forn, NCM, Forn_NCM) %>%
          summarise()

        NCM_Total = NCM

        NCM_mva = NCM
        NCM$Dif_Cod = NCM$Forn_NCM == NCM$NCM

        NCM = NCM[NCM$Dif_Cod == FALSE, ]

        NCM_Desc_Inc = data.frame()

        if (nrow(NCM) > 0) {
          primeiro = NCM %>%
            group_by(NCM) %>%
            summarise("Qtde" = length(NCM))
          colnames(primeiro) = c("NCM", "Qtde")

          segundo = NCM %>%
            group_by(Forn_NCM) %>%
            summarise("Qtde" = length(Forn_NCM))
          colnames(segundo) = c("NCM", "Qtde")

          NCM_Desc_Inc = rbind(primeiro, segundo)

          NCM_Desc_Inc = NCM_Desc_Inc %>%
            group_by(NCM) %>%
            summarise("Qtde" = length(NCM))

          NCM_Desc_Inc = merge(NCM_Desc_Inc,
                               base_NCM,
                               by.x = "NCM",
                               by.y = "NCM")

          NCM_Desc_Inc = NCM_Desc_Inc[, c(
            "NCM",
            "Categoria",
            "Descrição",
            "IPI",
            "Em.vigência",
            "Início.da.Vigência",
            "Fim.da.Vigência"
          )]
        }


        ########################################### CST

        base_dest = data.frame(
          "Ref_Prod" = str_trim(OrdemCompra$cd_Referencia),
          "Desc_Prod" = str_trim(OrdemCompra$ds_Prod),
          "OC_ICMS_orig" = str_trim(gsub(
            '[[:punct:]]', '', OrdemCompra$Sta_OrigemMat_PE
          )),
          "OC_ICMS_CST" = str_trim(gsub(
            '[[:punct:]]', '', OrdemCompra$TabTributacaoICMS
          )),
          "Ref_Forn" = str_trim(OrdemCompra$cdProdForn),
          "Faixa_ICMS" = str_trim(OrdemCompra$fx_ICMS_Saida),
          "Faixa_ICMS_Ali" = str_trim(OrdemCompra$Ali_ICMS_Saida)
        )

        #base_dest$CST_ICMS_Ali = ifelse(Forn_UF == 'SP' & base_dest$Faixa_ICMS == '4',
         #                                 '7%/12%',base_dest$CST_ICMS_Ali)

        base_dest = base_dest %>%
          group_by(Ref_Prod, Desc_Prod, OC_ICMS_orig, OC_ICMS_CST, Ref_Forn)

        base_dest = merge(base_dest, CST_NEO, by.x = "OC_ICMS_orig", by.y = "Proced")

        CST_IMPORTADO = base_dest[base_dest$OC_ICMS_orig == 'IMP',]

        teste_faixa_icms_DEST = base_dest
        teste_faixa_icms_DEST$Dif = ifelse((
          str_detect(teste_faixa_icms_DEST$Faixa_ICMS_Ali, "4") == T &
            str_detect(teste_faixa_icms_DEST$CST_ICMS_Ali, "4") == F
        )
        |
          (
            str_detect(teste_faixa_icms_DEST$Faixa_ICMS_Ali, "4") == F &
              str_detect(teste_faixa_icms_DEST$CST_ICMS_Ali, "4")
          ) == T,
        T,
        F)
        teste_faixa_icms_DEST = teste_faixa_icms_DEST[teste_faixa_icms_DEST$Dif ==
                                                        T, ]

        base_nf = data.frame(
          "Ref_Forn" = str_trim(PRODUTOS_NF$cProd),
          "NF_ICMS_orig" = str_trim(gsub(
            '[[:punct:]]', '', PRODUTOS_NF$ICMS_orig
          )),
          "NF_ICMS_CST" = str_trim(gsub(
            '[[:punct:]]', '', PRODUTOS_NF$ICMS_CST
          ))
        )

        base_nf = base_nf %>%
          group_by(Ref_Forn, NF_ICMS_orig, NF_ICMS_CST)

        base_nf = merge(base_nf, CST_NEO, by.x = "NF_ICMS_orig", by.y = "CST_ICMS")

        CST = merge(base_nf, base_dest, by.x = "Ref_Forn", by.y = "Ref_Forn")
        CST$Dif_Orig = CST$Proced == CST$OC_ICMS_orig
        CST$Dif_CST = CST$NF_ICMS_CST == CST$OC_ICMS_CST


        CST_Origem = CST[(CST$Dif_Orig == FALSE |
                            is.na(CST$Dif_CST) == T), ]

        if (nrow(CST_Origem) > 0) {
          primeiro = CST_Origem %>%
            group_by(Proced) %>%
            summarise("Qtde" = length(Proced))
          colnames(primeiro) = c("Origem", "Qtde")

          segundo = CST_Origem %>%
            group_by(OC_ICMS_orig) %>%
            summarise("Qtde" = length(OC_ICMS_orig))
          colnames(segundo) = c("Origem", "Qtde")

          CST_Origem_Desc = rbind(primeiro, segundo)

          CST_Origem_Desc = CST_Origem_Desc %>%
            group_by(Origem) %>%
            summarise("Qtde" = length(Origem))

          CST_Origem_Desc = merge(CST_Origem_Desc,
                                  CST_NEO,
                                  by.x = "Origem",
                                  by.y = "Proced")
        }


        CST_CST = CST[CST$Dif_CST == FALSE | is.na(CST$Dif_CST) == T, ]


        if (nrow(CST_CST) > 0) {
          primeiro = CST_CST %>%
            group_by(NF_ICMS_CST) %>%
            summarise("Qtde" = length(NF_ICMS_CST))
          colnames(primeiro) = c("Tributacao", "Qtde")

          segundo = CST_CST %>%
            group_by(OC_ICMS_CST) %>%
            summarise("Qtde" = length(OC_ICMS_CST))
          colnames(segundo) = c("Tributacao", "Qtde")

          CST_CST_Desc = rbind(primeiro, segundo)

          CST_CST_Desc = CST_CST_Desc %>%
            group_by(Tributacao) %>%
            summarise("Qtde" = length(Tributacao))

          CST_CST_Desc = merge(CST_CST_Desc,
                               CST_ICMS_TRI,
                               by.x = "Tributacao",
                               by.y = "CST_ICMS")
        }


        ########################################### UNIDADE DE MEDIDA

        base_teste_UM = base_teste_qtde_oc

        base_teste_UM$Dif_UM = base_teste_qtde_oc$UM_Prod_Forn == base_teste_qtde_oc$UM_Prod



        unidade_medida = base_teste_UM[base_teste_UM$Dif_UM == FALSE |
                                         is.na(base_teste_UM$Dif_UM) == T, ]

        unidade_medida_teste = unidade_medida
        unidade_medida_teste$chave_UM = paste(
          str_trim(unidade_medida_teste$UM_Prod),
          str_trim(unidade_medida_teste$UM_Prod_Forn),
          sep = "-"
        )

        unidade_medida_conv = merge(
          unidade_medida_teste,
          bd_UM_Conv,
          by.x = "chave_UM",
          by.y = "chave_UM",
          all.x = T
        )
        unidade_medida_conv = unidade_medida_conv[is.na(unidade_medida_conv$Fator_Conv), ]

        procurar_prod = paste("'", base_teste_UM$Ref_Prod, "'", sep = "")


        Compras_Realizadas = dbGetQuery(
          con,
          paste(
            "
    SELECT cd_Referencia,
    ds_Prod,cd_NotaFiscal,cd_UnidMed,
    Entradas_Notas.dt_EntregaMerc,
    Det_Entr_Notas_Fiscais.qtdItem
    FROM Entradas_Notas
    INNER JOIN Det_Entr_Notas_Fiscais WITH(NOLOCK) ON (Det_Entr_Notas_Fiscais.id_NotaFiscalEntrada = Entradas_Notas.id_NotaFiscalEntrada)
    INNER JOIN Materiais WITH(NOLOCK) ON (Materiais.id_Produto = Det_Entr_Notas_Fiscais.id_Produto)
    INNER JOIN Unid_Medidas WITH(NOLOCK) ON (Unid_Medidas.id_unidMed = Det_Entr_Notas_Fiscais.id_unidMed)
    WHERE Entradas_Notas.PierSitReg = 'ATV'
    AND DATEDIFF(month,Entradas_Notas.dt_EntregaMerc,GETDATE()) <= 6
    AND Entradas_Notas.StaNotaFiscalEnt = 'PRC'
    AND Entradas_Notas.TipoTransacaoCompra = 'CMI'
    AND Det_Entr_Notas_Fiscais.PierSitReg = 'ATV'
    AND Materiais.cd_Referencia IN (",
            paste(procurar_prod, collapse = ", "),
            ")",
            sep = ""
          )
        )


        Compras_Realizadas = Compras_Realizadas %>%
          group_by(cd_Referencia, ds_Prod, cd_UnidMed) %>%
          summarise(
            "Qtde_Itens" = sum(qtdItem),
            "Qtde_NF" = length(cd_Referencia),
            "Primeira_Compra" = min(dt_EntregaMerc),
            "Ultima_Compra" = max(dt_EntregaMerc),
            "Qtde_NF" = length(cd_Referencia)
          )

        Compras_Realizadas$MesesCompra = as.numeric(
          difftime(
            Compras_Realizadas$Ultima_Compra,
            Compras_Realizadas$Primeira_Compra,
            units = "days"
          )
        ) / 30
        Compras_Realizadas$QtdeMediaNF = Compras_Realizadas$Qtde_Itens / Compras_Realizadas$Qtde_NF
        Compras_Realizadas$QtdeMediaMeses = Compras_Realizadas$Qtde_Itens / Compras_Realizadas$MesesCompra





        ########################################### CFOP

        CFOP_Principal = ifelse(Forn_UF != Dest_UF, 6, 5)

        base_dest = data.frame(
          "Ref_Prod" = str_trim(OrdemCompra$cd_Referencia),
          "Desc_Prod" = str_trim(OrdemCompra$ds_Prod),
          "CFOP_Prod" = str_trim(gsub(
            '[[:punct:]]', '', OrdemCompra$Cod_CFOP
          )),
          "Ref_Forn" = str_trim(OrdemCompra$cdProdForn)
        )

        base_dest = base_dest %>%
          group_by(Ref_Prod, Desc_Prod, CFOP_Prod, Ref_Forn)

        base_dest$CFOP_Prod_Princ = str_sub(base_dest$CFOP_Prod, 1, 1)
        base_dest$CFOP_Prod_Outros = str_sub(base_dest$CFOP_Prod, 2, 4)

        base_nf = data.frame(
          "Ref_Forn" = str_trim(PRODUTOS_NF$cProd),
          "CFOP_Prod_Forn" = str_trim(PRODUTOS_NF$CFOP)
        )

        base_nf = base_nf %>%
          group_by(Ref_Forn, CFOP_Prod_Forn)

        base_nf$CFOP_Prod_Forn_Princ = str_sub(base_nf$CFOP_Prod_Forn, 1, 1)
        base_nf$CFOP_Prod_Forn_Outros = str_sub(base_nf$CFOP_Prod_Forn, 2, 4)
        base_nf$CFOP_Princ_Dif = ifelse(base_nf$CFOP_Prod_Forn_Princ != CFOP_Principal,
                                        FALSE,
                                        TRUE)

        CFOP_INC = merge(base_nf, base_dest, by.x = "Ref_Forn", by.y = "Ref_Forn")

        CFOP_teste = merge(base_nf, bdCFOP, by.x = "CFOP_Prod_Forn", by.y = "CFOP")
        CFOP_teste = merge(CFOP_teste,
                           PRODUTOS_NF,
                           by.x = "Ref_Forn",
                           by.y = "cProd")


        CFOP_Brinde = CFOP_teste[str_detect(str_to_upper(CFOP_teste$Descr_CFOP), "BRINDE") ==
                                   T, ]
        CFOP_Dev = CFOP_teste[str_detect(str_to_upper(CFOP_teste$Descr_CFOP), "DEVOLU") ==
                                T, ]
        CFOP_Anul = CFOP_teste[str_detect(str_to_upper(CFOP_teste$Descr_CFOP), "ANULA") ==
                                 T, ]
        CFOP_Trans = CFOP_teste[str_detect(str_to_upper(CFOP_teste$Descr_CFOP), "TRANSFER") ==
                                  T, ]
        CFOP_Amost = CFOP_teste[str_detect(str_to_upper(CFOP_teste$Descr_CFOP), "AMOSTRA") ==
                                  T, ]

        if (nrow(CFOP_INC) > 0) {
          CFOP_INC$Correto = CFOP_Principal
          CFOP_INC = CFOP_INC[CFOP_INC$CFOP_Princ_Dif == FALSE  |
                                is.na(CFOP_INC$CFOP_Princ_Dif) == T, ]
        } else{
          CFOP_INC = data.frame()
        }



        ########################################### conta contabil e consumo

        base_dest = ""

        base_dest = data.frame(
          "Ref_Prod" = str_trim(OrdemCompra$cd_Referencia),
          "Desc_Prod" = str_trim(OrdemCompra$ds_Prod),
          "CContabil" = str_trim(OrdemCompra$CContabil),
          "DSContabil" = str_trim(OrdemCompra$DSContabil),
          "CConsumo" = str_trim(OrdemCompra$CConsumo),
          "DSConsumo" = str_trim(OrdemCompra$DSConsumo)
        )

        base_dest = base_dest %>%
          group_by(Ref_Prod,
                   Desc_Prod,
                   CContabil,
                   DSContabil,
                   CConsumo,
                   DSConsumo) %>%
          summarise()

        base_dest$erro_ccconT = ifelse(Tipo_Tran == "CMI" &
                                         base_dest$CContabil != "1150401", T, F)
        base_dest$erro_ccconS = ifelse(Tipo_Tran == "CMI" &
                                         base_dest$CConsumo != "4110101", T, F)

        erro_ccconT = base_dest[base_dest$erro_ccconT == TRUE |
                                  is.na(base_dest$erro_ccconT) == T, ]
        erro_ccconS = base_dest[base_dest$erro_ccconS == TRUE |
                                  is.na(base_dest$erro_ccconS) == T, ]


        ########################################### ICMS

        base_dest = ""

        base_dest = data.frame(
          "Ref_Prod" = str_trim(OrdemCompra$cd_Referencia),
          "Desc_Prod" = str_trim(OrdemCompra$ds_Prod),
          "ICMS_Prod_Ali" = ifelse(
            OrdemCompra$Ali_Icms == 0 &
              OrdemCompra$VlrIcms > 0,
            round(
              OrdemCompra$VlrIcms / (OrdemCompra$PrecoUnitario * OrdemCompra$QtdSaldo) *
                100,
              2
            ),
            OrdemCompra$Ali_Icms
          ),
          "Ref_Forn" = str_trim(OrdemCompra$cdProdForn),
          "ICMS_DestVlr_Prod" = OrdemCompra$PrecoUnitario,
          "IPI_Prod_Ali" = ifelse(
            OrdemCompra$Ali_IPI == 0 &
              OrdemCompra$vlrIPI > 0,
            round(
              OrdemCompra$vlrIPI / (OrdemCompra$PrecoUnitario * OrdemCompra$QtdSaldo) *
                100,
              2
            ),
            OrdemCompra$Ali_IPI
          ),
          "OC_qtde" = OrdemCompra$QtdSaldo,
          "OC_UM" = OrdemCompra$cd_UnidMed,
          "ICMS_ST" = OrdemCompra$vlrICMSSubst,
          "ICMS_ST_MVA" = OrdemCompra$MVA_Aliq_MVA,
          "NCM_Prod" = str_trim(gsub('[[:punct:]]', '', OrdemCompra$NCM)),
          "Sta_OrigemMat_PE" = str_trim(OrdemCompra$Sta_OrigemMat_PE)
        )

        base_dest = base_dest %>%
          group_by(Ref_Prod, Desc_Prod, Ref_Forn, OC_UM,NCM_Prod,Sta_OrigemMat_PE) %>%
          summarise(
            "ICMS_Prod_Ali" = sum(ICMS_Prod_Ali) / length(Ref_Prod),
            "ICMS_DestVlr_Prod" = sum(ICMS_DestVlr_Prod) / length(Ref_Prod),
            "IPI_Prod_Ali" = sum(IPI_Prod_Ali) / length(Ref_Prod),
            "OC_qtde" = sum(OC_qtde),
            "ICMS_ST" = sum(ICMS_ST),
            "ICMS_ST_MVA" = sum(ICMS_ST_MVA) / length(Ref_Prod)
          )


        prod_nao_relacionados_base = base_dest[is.na(base_dest$Ref_Forn) == T, ]

        base_dest = base_dest[is.na(base_dest$Ref_Forn) == F, ]

        base_nf = ""

        base_nf = data.frame(
          "Ref_Forn" = str_trim(PRODUTOS_NF$cProd),
          "ICMS_Forn_Ali" = str_trim(PRODUTOS_NF$ICMS_pICMS),
          "ICMS_Forn_Vlr" = str_trim(PRODUTOS_NF$ICMS_vICMS),
          "ICMS_QtdeForn_Prod" = str_trim(PRODUTOS_NF$qCom),
          "ICMS_FornVlr_Prod" = str_trim(PRODUTOS_NF$vUnCom),
          "ICMS_FornVlr_Prod_Total" = str_trim(PRODUTOS_NF$vProd),
          "IPI_Forn_Ali" = str_trim(PRODUTOS_NF$IPI_pIPI),
          "IPI_Forn_Vlr" = str_trim(PRODUTOS_NF$IPI_vIPI),
          "Forn_UM" = f.sem_ascento(str_trim(PRODUTOS_NF$uCom)),
          "ICMS_vBCST" = str_trim(PRODUTOS_NF$ICMS_vBCST),
          "ICMS_pICMSST" = str_trim(PRODUTOS_NF$ICMS_pICMSST),
          "ICMS_vICMSST" = str_trim(PRODUTOS_NF$ICMS_vICMSST),
          "NCM_Forn" = str_trim(PRODUTOS_NF$NCM),
          "ICMS_orig" = str_trim(PRODUTOS_NF$ICMS_orig)
        )

        base_nf = merge(base_nf,CST_NEO,by.x = "ICMS_orig",by.y = "CST_ICMS",all.x = T)
        base_nf$ICMS_Orig_Forn =  base_nf$Proced
        base_nf$Nacionalidade = NULL
        base_nf$Proced = NULL
        base_nf$CST_ICMS_Ali = NULL
        base_nf$CST_ICMS_Desc = NULL

        base_nf = base_nf %>%
          group_by(Ref_Forn, Forn_UM,NCM_Forn,ICMS_Orig_Forn) %>%
          summarise(
            "ICMS_Forn_Ali" = sum(as.numeric(ICMS_Forn_Ali)) / length(Ref_Forn),
            "ICMS_Forn_Vlr" = sum(as.numeric(ICMS_Forn_Vlr)),
            "ICMS_QtdeForn_Prod" = sum(as.numeric(ICMS_QtdeForn_Prod)),
            "ICMS_FornVlr_Prod" = sum(as.numeric(ICMS_FornVlr_Prod)) / length(Ref_Forn),
            "ICMS_FornVlr_Prod_Total" = sum(as.numeric(ICMS_FornVlr_Prod_Total)),
            "IPI_Forn_Ali" = sum(as.numeric(IPI_Forn_Ali)) / length(Ref_Forn),
            "IPI_Forn_Vlr" = sum(as.numeric(IPI_Forn_Vlr)),
            "ICMS_vBCST" = sum(as.numeric(ICMS_vBCST)),
            "ICMS_pICMSST" = sum(as.numeric(ICMS_pICMSST)) / length(Ref_Forn),
            "ICMS_vICMSST" = sum(as.numeric(ICMS_vICMSST))
          )

        base_nf[is.na(base_nf)] = 0

        base_teste_icms_ali = merge(base_nf, base_dest)
        base_teste_icms_ali$chave_UM = paste(
          str_trim(base_teste_icms_ali$Forn_UM),
          str_trim(base_teste_icms_ali$OC_UM),
          sep = "-"
        )

        if (nrow(base_teste_icms_ali) > 0) {
          valores_UM = merge(
            base_teste_icms_ali,
            bd_UM_Conv,
            by.x = "chave_UM",
            by.y = "chave_UM",
            all.x = T
          )

          valores_UM$Fator_Conv_UM = ifelse(
            is.na(valores_UM$Fator_Conv) == TRUE |
              valores_UM$Fator_Conv == 'NA',
            1,
            ifelse(
              valores_UM$UM_EMPRESA == valores_UM$UM_FORN,
              1,
              valores_UM$Fator_Conv
            )
          )
          valores_UM$Mult_Div_UM = ifelse(
            is.na(valores_UM$Fator_Conv) == TRUE |
              valores_UM$Fator_Conv == 'NA',
            "MUL",
            ifelse(
              valores_UM$UM_EMPRESA == valores_UM$UM_FORN,
              "MUL",
              valores_UM$OperMatConvUnidMed
            )
          )

          tem_conv_UM = valores_UM[str_trim(valores_UM$Forn_UM) != str_trim(valores_UM$OC_UM), ]

          if (nrow(tem_conv_UM) > 0) {
            tem_conv_UM = tem_conv_UM %>%
              group_by(chave_UM, Fator_Conv_UM, Mult_Div_UM) %>%
              summarise()

          }


          valores_UM = valores_UM[, c("Ref_Prod", "Fator_Conv_UM", "Mult_Div_UM")]

          base_teste_icms_ali[is.na(base_teste_icms_ali)] = 0

          base_teste_icms_ali = merge(
            base_teste_icms_ali,
            valores_UM,
            by.x = "Ref_Prod",
            by.y = "Ref_Prod",
            all.x = T
          )

          base_teste_icms_ali$nao_tem_UM = ifelse((
            is.na(base_teste_icms_ali$Fator_Conv) == TRUE |
              base_teste_icms_ali$Fator_Conv == 'NA'
          )
          &
            base_teste_icms_ali$Forn_UM != base_teste_icms_ali$OC_UM,
          TRUE,
          FALSE
          )

          nao_tem_UM = base_teste_icms_ali[base_teste_icms_ali$base_teste_icms_ali == TRUE, ]

          base_teste_icms_ali$Fator_Conv_UM = ifelse(
            is.na(base_teste_icms_ali$Fator_Conv) == TRUE |
              base_teste_icms_ali$Fator_Conv == 'NA',
            1,
            base_teste_icms_ali$Fator_Conv_UM
          )

          base_teste_icms_ali$Mult_Div_UM = ifelse(
            is.na(base_teste_icms_ali$Fator_Conv) == TRUE |
              base_teste_icms_ali$Fator_Conv == 'NA',
            "DIV",
            base_teste_icms_ali$Mult_Div_UM
          )

          base_teste_icms_ali$ICMS_QtdeForn_Prod = ifelse(
            base_teste_icms_ali$Mult_Div_UM == 'DIV',
            base_teste_icms_ali$ICMS_QtdeForn_Prod * base_teste_icms_ali$Fator_Conv_UM,
            base_teste_icms_ali$ICMS_QtdeForn_Prod / base_teste_icms_ali$Fator_Conv_UM
          )

          base_teste_icms_ali$ICMS_FornVlr_Prod = ifelse(
            base_teste_icms_ali$Mult_Div_UM == 'DIV',
            base_teste_icms_ali$ICMS_FornVlr_Prod / base_teste_icms_ali$Fator_Conv_UM,
            base_teste_icms_ali$ICMS_FornVlr_Prod * base_teste_icms_ali$Fator_Conv_UM
          )

          base_teste_icms_ali$Dif_qtde = base_teste_icms_ali$ICMS_QtdeForn_Prod  - base_teste_icms_ali$OC_qtde




          base_teste_icms_ali$VlrProd_base_nf = base_teste_icms_ali$ICMS_QtdeForn_Prod * base_teste_icms_ali$ICMS_DestVlr_Prod
          base_teste_icms_ali$VlrProd_Dif = base_teste_icms_ali$VlrProd_base_nf - base_teste_icms_ali$ICMS_FornVlr_Prod_Total
          base_teste_icms_ali$VlrProd_Dif_perc = (base_teste_icms_ali$VlrProd_Dif / base_teste_icms_ali$VlrProd_base_nf) * 100
          base_teste_icms_ali$VlrProd_Dif_an = ifelse(abs(base_teste_icms_ali$VlrProd_Dif_perc) > 0.5,
                                                      FALSE,
                                                      TRUE)
          base_teste_icms_ali$VlrProd_UNIT_Dif = base_teste_icms_ali$ICMS_FornVlr_Prod - base_teste_icms_ali$ICMS_DestVlr_Prod
          base_teste_icms_ali$VlrProd_UNIT_Dif_per = (
            base_teste_icms_ali$VlrProd_UNIT_Dif / base_teste_icms_ali$ICMS_DestVlr_Prod
          ) * 100
          base_teste_icms_ali$VlrProd_UNIT_Dif_per_an = ifelse(abs(base_teste_icms_ali$VlrProd_UNIT_Dif_per) > 0.5,
                                                               FALSE,
                                                               TRUE)

          base_teste_icms_ali$ICMS_Dif_Ali = str_trim(base_teste_icms_ali$ICMS_Forn_Ali) == str_trim(base_teste_icms_ali$ICMS_Prod_Ali)
          base_teste_icms_ali$ICMS_base_nf = base_teste_icms_ali$VlrProd_base_nf * (base_teste_icms_ali$ICMS_Prod_Ali /
                                                                                      100)
          base_teste_icms_ali$ICMS_dif_vlr = base_teste_icms_ali$ICMS_base_nf - base_teste_icms_ali$ICMS_Forn_Vlr
          base_teste_icms_ali$icms_dif_per = abs(base_teste_icms_ali$ICMS_dif_vlr / base_teste_icms_ali$ICMS_base_nf) * 100
          base_teste_icms_ali$icms_dif_per_an = ifelse(abs(base_teste_icms_ali$icms_dif_per) > 0.5,
                                                       FALSE,
                                                       TRUE)

          icms_inc_ali = base_teste_icms_ali[base_teste_icms_ali$ICMS_Dif_Ali == FALSE, ]
          icms_inc_vlr = base_teste_icms_ali[abs(base_teste_icms_ali$icms_dif_per) > 0.5 &
                                               abs(base_teste_icms_ali$ICMS_dif_vlr) > 0, ]

          base_teste_icms_ali$IPI_Dif_Ali = str_trim(base_teste_icms_ali$IPI_Forn_Ali) == str_trim(base_teste_icms_ali$IPI_Prod_Ali)
          base_teste_icms_ali$IPI_base_nf = base_teste_icms_ali$VlrProd_base_nf * (base_teste_icms_ali$IPI_Prod_Ali /
                                                                                     100)
          base_teste_icms_ali$IPI_dif_vlr = base_teste_icms_ali$IPI_base_nf - base_teste_icms_ali$IPI_Forn_Vlr
          base_teste_icms_ali$IPI_dif_per = abs(base_teste_icms_ali$IPI_dif_vlr / base_teste_icms_ali$IPI_base_nf) * 100
          base_teste_icms_ali$IPI_dif_per_an = ifelse(abs(base_teste_icms_ali$IPI_dif_per) > 0.5,
                                                      FALSE,
                                                      TRUE)

          ipi_inc_ali = base_teste_icms_ali[base_teste_icms_ali$IPI_Dif_Ali == FALSE, ]
          ipi_inc_vlr = base_teste_icms_ali[abs(base_teste_icms_ali$IPI_dif_per) > 0.5 &
                                              abs(base_teste_icms_ali$IPI_dif_vlr) > 0, ]

          preco_unitario = base_teste_icms_ali[base_teste_icms_ali$VlrProd_UNIT_Dif_per_an == FALSE, ]
          valor_total_produto = base_teste_icms_ali[base_teste_icms_ali$VlrProd_Dif_an == FALSE, ]


          base_teste_icms_ali$ICMS_ST_calc = (base_teste_icms_ali$VlrProd_base_nf * (1+(base_teste_icms_ali$ICMS_ST_MVA /
                                                                                          100))*(0.18))-(base_teste_icms_ali$VlrProd_base_nf*(base_teste_icms_ali$ICMS_Prod_Ali/100))

          #base_teste_icms_ali$ICMS_ST_calc = base_teste_icms_ali$ICMS_ST_MVA /
          #  100 * base_teste_icms_ali$ICMS_base_nf


          base_teste_icms_ali$ICMS_ST_dif = base_teste_icms_ali$ICMS_vICMSST - base_teste_icms_ali$ICMS_ST_calc
          base_teste_icms_ali$ST_Forn = ifelse(base_teste_icms_ali$ICMS_vICMSST >
                                                 0.01, 1, 0)
          base_teste_icms_ali$ST_OC = ifelse(base_teste_icms_ali$ICMS_ST > 0.01, 1, 0)

          base_teste_icms_ali$MVA_Forn = (
            round((base_teste_icms_ali$ICMS_vBCST) / (base_teste_icms_ali$ICMS_FornVlr_Prod_Total+base_teste_icms_ali$IPI_Forn_Vlr) - 1,2)
          ) * 100
          icms_ST_dif = base_teste_icms_ali[base_teste_icms_ali$ST_Forn + base_teste_icms_ali$ST_OC > 0, ]
          icms_ST_dif = icms_ST_dif[abs(base_teste_icms_ali$ICMS_ST_dif) > 0, ]

          icms_ST_dif = merge(icms_ST_dif,
                              NCM_mva,
                              by.x = "Ref_Prod",
                              by.y = "Ref_Prod")

          icms_ST_dif = icms_ST_dif[abs(icms_ST_dif$ICMS_ST_dif) > 1, ]

          icms_ST_dif_mva = icms_ST_dif[icms_ST_dif$MVA_Forn != icms_ST_dif$ICMS_ST_MVA, ]



          qtde_rec_acima = nrow(base_teste_icms_ali[base_teste_icms_ali$Dif_qtde >
                                                      0, ])
          qtde_rec_abaixo = nrow(base_teste_icms_ali[base_teste_icms_ali$Dif_qtde <
                                                       0, ])
          qtde_rec_exata = nrow(base_teste_icms_ali[base_teste_icms_ali$Dif_qtde ==
                                                      0, ])


          base_rec_acima = base_teste_icms_ali[base_teste_icms_ali$Dif_qtde >
                                                 0, ]

          Compras_Realizadas =   Compras_Realizadas %>%
            filter(cd_Referencia %in% as.list(base_rec_acima$Ref_Prod))

          Compras_Realizadas = as.data.frame(Compras_Realizadas)


          # TESTE IPI NCM ----------------------------------------------------------------------------




          NCM_IPI_Produtos = base_teste_icms_ali %>%
            group_by(Ref_Prod, Desc_Prod, Ref_Forn) %>%
            summarise(
              "OC IPI" = sum(IPI_Prod_Ali) / length(Ref_Prod),
              "NF IPI" = sum(IPI_Forn_Ali) / length(Ref_Prod)
            )


          NCM_IPI_Produtos = merge(
            NCM_IPI_Produtos,
            NCM_Total,
            by.x = "Ref_Prod",
            by.y = "Ref_Prod",
            all.x = T
          )
          NCM_IPI_Produtos = NCM_IPI_Produtos[, c("Ref_Prod",
                                                  "Desc_Prod.x",
                                                  "Ref_Forn.x",
                                                  "OC IPI",
                                                  "NF IPI",
                                                  "Forn_NCM",
                                                  "NCM")]
          colnames(NCM_IPI_Produtos) = c("Ref_Prod",
                                         "Desc_Prod",
                                         "Ref_Forn",
                                         "OC IPI",
                                         "NF IPI",
                                         "Forn_NCM",
                                         "NCM")

          NCM_IPI_OC = merge(
            NCM_IPI_Produtos,
            base_NCM,
            by.x = "NCM",
            by.y = "NCM",
            all.x = T
          )
          NCM_IPI_OC = NCM_IPI_OC[, c("Ref_Prod",
                                      "Desc_Prod",
                                      "Ref_Forn",
                                      "OC IPI",
                                      "NCM",
                                      "IPI")]
          NCM_IPI_OC = NCM_IPI_OC[NCM_IPI_OC$`OC IPI` != NCM_IPI_OC$`IPI`, ]

          NCM_IPI_NF = merge(
            NCM_IPI_Produtos,
            base_NCM,
            by.x = "Forn_NCM",
            by.y = "NCM",
            all.x = T
          )
          NCM_IPI_NF = NCM_IPI_NF[, c("Ref_Prod",
                                      "Desc_Prod",
                                      "Ref_Forn",
                                      "NF IPI",
                                      "Forn_NCM",
                                      "IPI")]
          NCM_IPI_NF = NCM_IPI_NF[NCM_IPI_NF$`NF IPI` != NCM_IPI_NF$`IPI`, ]


        } else {
          valores_UM = data.frame()
          tem_conv_UM = data.frame()
          base_teste_icms_ali = data.frame()
          NCM_IPI_NF = data.frame()
          NCM_IPI_OC = data.frame()
          NCM_IPI_Produtos = data.frame()
          Compras_Realizadas = data.frame()

          qtde_rec_acima = 0
          qtde_rec_abaixo = 0
          qtde_rec_exata = 0


          base_rec_acima = data.frame()

          icms_inc_ali = data.frame()
          icms_inc_vlr = data.frame()

          ipi_inc_ali = data.frame()
          ipi_inc_vlr = data.frame()

          preco_unitario = data.frame()
          valor_total_produto = data.frame()

          icms_ST_dif = data.frame()
          icms_ST_dif = data.frame()

          icms_ST_dif = data.frame()

          icms_ST_dif_mva = data.frame()

          base_rec_acima = data.frame()

          nao_tem_UM = data.frame()
        }


        ############################### RESUMO

        RESUMO = data.frame(
          "Análise" = c(
            "01 - CNPJ (NF X OC)",
            "02 - MERCADORIA EXISTENTE NA NOTA FISCAL, PORÉM SEM RELACIONAMENTO CADASTRADO",
            "03 - MERCADORIA CADASTRADA NO ARECO, PORÉM SEM RELACIONAMENTO COM O FORNECEDOR",
            "04 - AUTORIZAÇÃO DE USO NF-e",
            "05 - CÓDIGO DE PRODUTO USADO PARA MAIS DE ITEM NA NF",
            "06 - UNIDADE DE MEDIDA (NF X OC)",
            "07 - UNIDADE DE MEDIDA - CONVERSÃO INEXISTENTE",
            "08 - ERROS NA CONTA CONTÁBIL - MATERIAL",
            "09 - ERROS NA CONTA CONSUMO - MATERIAL",
            "10 - QUANTIDADE DE MATERIAL DA NF SUPERIOR AO DA OC",
            "11 - TIPO TRANSAÇÃO X HISTÓRICO FINANCEIRO",
            "12 - FRETE",
            "13 - CONDIÇÃO DE PAGAMENTO - QUANTIDADE DE PARCELAS",
            "14 - CONDIÇÃO DE PAGAMENTO - PRAZO MÉDIO",
            "15 - PREÇO UNITÁRIO",
            "16 - VALOR TOTAL DO PRODUTO",
            "17 - VALOR FINANCEIRO X VALOR DA NF",
            "18 - NCM",
            "19 - CST/ICMS - ORIGEM CADASTRADA COMO IMPORTADA",
            "20 - CST/ICMS - ORIGEM",
            "21 - CST/ICMS - TRIBUTAÇÃO",
            "22 - CST/ICMS - ALIQ FAIXA X ALI CST (CADASTRO DO PRODUTO)",
            "23 - CFOP PRINCIPAL",
            "24 - CFOP BRINDE, AMOSTRA, ANULAÇÃO, DEVOLUÇÃO e TRANSFERÊNCIA",
            "25 - ICMS - ALÍQUOTA",
            "26 - ICMS - VALOR",
            "27 - ICMS ST - MVA",
            "28 - ICMS ST - VALOR",
            "29 - IPI - ALÍQUOTA",
            "30 - IPI - VALOR",
            "31 - IPI - NCM (OC)",
            "32 - IPI - NCM (NF)"
          ),
          "Resultado" = c(
            OC_CNPJ != Forn_CNPJ,
            qtde_nao_relacionados > 0,
            nrow(prod_nao_relacionados_base) > 0,
            DG_USO != "Autorizado o uso da NF-e",
            nrow(ref_qtde_acima1) > 0,
            nrow(unidade_medida) > 0,
            nrow(nao_tem_UM) > 0,
            nrow(erro_ccconT) > 0,
            nrow(erro_ccconS) > 0,
            nrow(Compras_Realizadas) > 0,
            Tipo_Tran == 'CMI' &
              cd_HistFin != '1300',
            frete_OC != Trans_Frete,
            Qtde_ParcelasNF != QtdeParcedlasOC |
              Pagto_nao_dest == 1,
            PrazoMedioNF < PrazoMedioOC |
              Pagto_nao_dest == 1,
            nrow(preco_unitario) > 0,
            nrow(valor_total_produto) > 0,
            abs(as.numeric(VlrFIN_NF) - as.numeric(vNF_Total)) >
              1,
            nrow(NCM) > 0,
            nrow(CST_IMPORTADO) > 0,
            nrow(CST_Origem) > 0,
            nrow(CST_CST) > 0,
            nrow(teste_faixa_icms_DEST) > 0,
            nrow(CFOP_INC) > 0,
            nrow(CFOP_Brinde) > 0 |
              nrow(CFOP_Dev) > 0 |
              nrow(CFOP_Anul) > 0 |
              nrow(CFOP_Trans) > 0 | nrow(CFOP_Amost) > 0,
            nrow(icms_inc_ali) > 0,
            nrow(icms_inc_vlr) > 0,
            nrow(icms_ST_dif_mva) > 0,
            nrow(icms_ST_dif) > 0,
            nrow(ipi_inc_ali) > 0,
            nrow(ipi_inc_vlr) > 0,
            nrow(NCM_IPI_OC) > 0,
            nrow(NCM_IPI_NF) > 0
          )
        )

        if (nrow(base_teste_icms_ali) == 0) {
          RESUMO$Resultado = TRUE

        }


        message("SALVANDO RELATÓRIOS NA PASTA RECEBIMENTO_DOCS")

        arquivo_salvar_GLOG = paste(pasta_GLOG,
                                    "NF_",
                                    DG_NF,
                                    "_OC_",
                                    OC,
                                    "_CDFORN_",
                                    cd_Forn_Areco,
                                    "_GLOG",
                                    sep = "")
        arquivo_salvar_FISCAL = paste(pasta_GLOG,
                                      "NF_",
                                      DG_NF,
                                      "_OC_",
                                      OC,
                                      "_CDFORN_",
                                      cd_Forn_Areco,
                                      "_FISCAL",
                                      sep = "")


        render(
          "J:/Suporte/DadosR/Conf_Mercadoria.rmd",
          output_file = arquivo_salvar_GLOG
        )

        render(
          "J:/Suporte/DadosR/Conf_Mercadoria_Fiscal.rmd",
          output_file = arquivo_salvar_FISCAL
        )


        message("SALVANDO HISTÓRICO - BI")

        if (nrow(base_teste_icms_ali) > 0) {
          dados_analise_entrada_arq = read.csv2(
            paste(pasta_SUPORTE, "Dados_Entrada.csv", sep = ""),
            header = TRUE,
            sep = ";",
            dec = ",",
            fill = TRUE,
            na.strings = "N/A"
          )


          ID_N = max(dados_analise_entrada_arq$ID)

          adicionar_dados_entrada = data.frame(
            ID = 1 + ID_N,
            NF = DG_NF,
            cd_Forn_Areco = cd_Forn_Areco,
            DG_EMISSAO = DG_EMISSAO,
            Forn_RazaoSocial = Forn_RazaoSocial,
            OC = OC
          )

          adicionar_dados_PRODUTO = merge(
            base_teste_UM,
            tem_SC,
            all.x = T,
            by.x = "Ref_Prod",
            by.y = "cd_Referencia"
          )

          adicionar_dados_todos = merge(adicionar_dados_entrada, adicionar_dados_PRODUTO)

          dados_analise_entrada = adicionar_dados_todos[, c(
            "ID",
            "NF",
            "DG_EMISSAO",
            "OC",
            "cd_Forn_Areco",
            "Forn_RazaoSocial",
            "Ref_Prod",
            "NF_QtdeProd.x",
            "cd_Sc",
            "NomeSolicitante",
            "Quantidade"
          )]

          for (i in 1:nrow(dados_analise_entrada)) {
            ID_N = ID_N + 1
            dados_analise_entrada$ID[i] = ID_N
          }

          dados_analise_entrada$NomeSolicitante = ifelse(
            is.na(dados_analise_entrada$NomeSolicitante) == TRUE,
            "Nenhum",
            dados_analise_entrada$NomeSolicitante
          )
          dados_analise_entrada$Quantidade = ifelse(
            is.na(dados_analise_entrada$Quantidade) == TRUE,
            0,
            dados_analise_entrada$Quantidade
          )

          dados_analise_entrada_arq = rbind(dados_analise_entrada_arq, dados_analise_entrada)

          dados_analise_entrada_arq = dados_analise_entrada_arq[order(-dados_analise_entrada_arq$ID), ]

          dados_analise_entrada_arq$NomeSolicitante = str_trim(dados_analise_entrada_arq$NomeSolicitante)

          write.table(
            dados_analise_entrada_arq,
            file = paste(pasta_SUPORTE, "Dados_Entrada.csv", sep = ""),
            sep = ";",
            na = "",
            quote = TRUE,
            row.names = FALSE
          )

          write.table(
            dados_analise_entrada_arq,
            file = paste("J:/Suporte/DadosR/", "Dados_Entrada.csv", sep = ""),
            sep = ";",
            na = "",
            quote = TRUE,
            row.names = FALSE
          )


        }

        if (envio_email == "S" & qtde_nao_relacionados == 0) {

          RESUMO_COMPRAS = RESUMO[1:17,]

          RESUMO_COMPRAS = RESUMO_COMPRAS[RESUMO_COMPRAS$Resultado==T,]

          RESUMO_FISCAL = RESUMO[RESUMO$Resultado==T,]


          EMAIL_FUN = data.frame(
            "Código" = c("1034",
                         "11474",
                         "920",
                         "919"),
            "Descrição" = c("SUZI",
                            "NEIDE",
                            "YOLANDA",
                            "MARCO"),
            "email" = c(
              "suzi.lima@neotermica.com.br",
              "neide.araujo@neotermica.com.br",
              "yolanda@neotermica.com.br",
              "marco.martins@neotermica.com.br"
            )
          )

          arquivo_salvar_EMAIL = paste(pasta_SUPORTE, "Conf_Mercadoria_Email", sep =
                                         "")


          cod_solicitante_para_email = tem_SC %>%
            group_by(
              "cdSolicitante" = str_trim(cdSolicitante),
              "NomeSolicitante" = str_trim(NomeSolicitante)
            ) %>%
            summarise("Qtde" = length(cdSolicitante))

          obs_SC_email = tem_SC %>%
            group_by(cd_Sc, Observacao) %>%
            summarise("Qtde" = length(cdSolicitante))



          email_tabela = dados_analise_entrada
          email_tabela$NomeSolicitante = str_trim(email_tabela$NomeSolicitante)

          email_tabela_qtde = email_tabela %>%
            group_by(NomeSolicitante) %>%
            summarise("Qtde" = length(NomeSolicitante))

          email_tabela_qtde = email_tabela_qtde[email_tabela_qtde$NomeSolicitante != "Nenhum", ]

          email_qtde = nrow(email_tabela_qtde)

          if (email_qtde > 0) {
            message("ENVIANDO E-MAIL À EQUIPE COMERCIAL!")

            email_tabela_enviar = data.frame()

            for (i in 1:email_qtde) {
              email_tabela_enviar = email_tabela[email_tabela$NomeSolicitante == email_tabela_qtde$NomeSolicitante[i], ]

              endereco_email_inter = merge(
                email_tabela_enviar,
                cod_solicitante_para_email,
                by.x = "NomeSolicitante",
                by.y = "NomeSolicitante"
              )

              endereco_email = merge(endereco_email_inter,
                                     EMAIL_FUN,
                                     by.x = "cdSolicitante",
                                     by.y = "Código")
              destinatario_email = max(endereco_email$NomeSolicitante)
              forn_email = max(endereco_email$Forn_RazaoSocial)
              NF_email = max(endereco_email$NF)
              SC_email = max(endereco_email$cd_Sc)
              endereco_email = max(endereco_email$email)

              email_tabela_enviar = merge(email_tabela_enviar,
                                          base_teste_UM,
                                          by.x = "Ref_Prod",
                                          by.y = "Ref_Prod")
              email_tabela_enviar = merge(email_tabela_enviar,
                                          obs_SC_email,
                                          by.x = "cd_Sc",
                                          by.y = "cd_Sc")
              email_tabela_enviar = email_tabela_enviar[, c(
                "cd_Sc",
                "Ref_Prod",
                "Desc_Prod",
                "Quantidade",
                "NF_QtdeProd.x",
                "Dif_qtde",
                "Observacao"
              )]
              email_tabela_enviar$Dif_qtde = email_tabela_enviar$NF_QtdeProd.x - email_tabela_enviar$Quantidade


              message(paste("ENVIANDO E-MAIL PARA ", destinatario_email, sep =
                              ""))

              titulo_email = paste("NEOTÉRMICA - Solicitação de Compra em Recebimento - NF ",
                                   NF_email,
                                   sep = "")


              email <- render_email('Conf_Mercadoria_Email.Rmd')

              email %>%
                smtp_send(
                  from = "vend1neotermica@gmail.com",
                  to = endereco_email,
                  subject = titulo_email,
                  credentials = creds_key(id = "gmail2")
                )

            }

          } else {
            message("NENHUM E-MAIL FOI ENVIADO À EQUIPE COMERCIAL!")

          }

          if (email_qtde > 0) {
            message("E-MAILS ENVIADOS À EQUIPE COMERCIAL!")

          }

          message("ENVIANDO E-MAIL PARA CONHECIMENTO - DIR!")

          email <- render_email('Conf_Mercadoria_EmailConhecimento.Rmd')

          email %>%
            smtp_send(
              from = "vend1neotermica@gmail.com",
              to = email_DIR,
              subject = paste("Emissão | RecM - NF ", DG_NF, " | ", Forn_RazaoSocial, sep =
                                ""),
              credentials = creds_key(id = "gmail2")
            )


          message("ENVIANDO E-MAIL PARA CONHECIMENTO - COMPRAS!")

          email <- render_email('Conf_Mercadoria_Email_COMPRAS.Rmd')

          email %>%
            smtp_send(
              from = "vend1neotermica@gmail.com",
              to = email_COM,
              subject = paste("Emissão | RecM - NF ", DG_NF, " | ", Forn_RazaoSocial, sep =
                                ""),
              credentials = creds_key(id = "gmail2")
            )


          message("ENVIANDO E-MAIL PARA CONHECIMENTO - FISCAL e LOGÍSTICA!")

          email <- render_email('Conf_Mercadoria_Email_FISCAL.Rmd')

          email %>%
            smtp_send(
              from = "vend1neotermica@gmail.com",
              to = c(email_FIS,email_LOG),
              subject = paste("Emissão | RecM - NF ", DG_NF, " | ", Forn_RazaoSocial, sep =
                                ""),
              credentials = creds_key(id = "gmail2")
            )


          #create_smtp_creds_key(
          #  id = "gmail2",
          #  user = "vend1neotermica@gmail.com",
          #  host = "smtp.gmail.com",
          #  port = 465,
          #  use_ssl = TRUE
          #)
          #rumpxchwrnesdtxp





        }

        if (qtde_nao_relacionados > 0) {
          message(
            "EXISTEM PRODUTOS NÃO RELACIONADOS. É NECESSÁRIO QUE REALIZE O RELACIONAMENTO PARA QUE OS E-MAILS POSSAM SER ENVIADOS À EQUIPE COMERCIAL!"
          )

          message("ENVIANDO E-MAIL COM O ERRO DE RELACIONAMENTO.")

          lote_dados_final$Prob_XML[var_consulta] = "NÃO RELACIONADO"
          lote_dados_final$Prob_OC[var_consulta] = "NÃO RELACIONADO"

          email <- render_email('Conf_Mercadoria_EmailERRO.Rmd')

          email %>%
            smtp_send(
              from = "vend1neotermica@gmail.com",
              to = c(email_FIS,email_LOG,email_DIR,email_COM),
              subject = paste("ERRO - Itens Não Relacionados - Emissão | RecM - NF ", DG_NF, " | ", Forn_RazaoSocial, sep =
                                ""),
              credentials = creds_key(id = "gmail2")
            )

        }

        message(paste(
          "!!! FINALIZADO COM SUCESSO !!! CHAVE: ",
          XML_NF,
          " | OC: ",
          OC,
          sep = ""
        ))

      } else {
        message(paste(
          "!!! EXECUÇÃO NÃO REALIZADA !!! CHAVE: ",
          XML_NF,
          " | OC: ",
          OC,
          sep = ""
        ))
      }


    } else {
      message(paste(
        "!!! EXECUÇÃO NÃO REALIZADA !!! CHAVE: ",
        XML_NF,
        " | OC: ",
        OC,
        sep = ""
      ))
    }

    nao_rodou = 1

  }, error = function(x) {
      message(paste("!!! ERRO !!! ANÁLISE NÃO FINALIZADA !!!", XML_NF,"OC",OC, sep = " - "))
    })

  if (nao_rodou==0){
    lote_dados_final$Prob_XML[var_consulta] = "NÃO RODOU"
    lote_dados_final$Prob_OC[var_consulta] = "NÃO RODOU"
  }
}

if (nrow(lote_dados_final) > 0) {

  lote_dados_final$Prob_XML = ifelse(lote_dados_final$Prob_XML=="","NÃO RODOU",lote_dados_final$Prob_XML)
  lote_dados_final$Prob_OC = ifelse(lote_dados_final$Prob_OC=="","NÃO RODOU",lote_dados_final$Prob_OC)
  message("RESUMO")
  print(lote_dados_final)
}

rstudioapi::showDialog ( "Finalizado" ,  "Resultados Gerados na pasta J:/Recebimento_Docs!" ,  url  =  "" )

}




