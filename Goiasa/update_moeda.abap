*----------------------------------------------------------------------*
* Report ZFI_COTACAO_MOEDAS                                            *
*----------------------------------------------------------------------*
* Autor......: MARCUS VINÍCIUS MACIEL VIEIRA                           *
* Data.......: 12/07/2024                                              *
* Descrição  : Automação de atualização na tabela de moeda estrangeira *
* Transação..: ZFI00040                                                *
*----------------------------------------------------------------------*
*                     Histórico das modificações                       *
*----------------------------------------------------------------------*
*   Data     |  Nome       | Request    | Descrição                    *
*----------------------------------------------------------------------*
* 26/06/2024 | mvmaciell  | ECDK998509 | Desenvolvimento inicial       *
* 12/07/2024 | mvmaciell  | ECDK998509 | Ajustes no envia email/http   *
* 30/07/2024 | mvmaciell  | ECDK998510 | Inclusão de novas cotações    *
*----------------------------------------------------------------------*
* Programa de automação de atualização na tabela de moeda estrageira   *
*----------------------------------------------------------------------*
REPORT ZFI_COTACAO_MOEDAS.

TABLES: adr6.

SELECTION-SCREEN BEGIN OF BLOCK a3 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_data_i TYPE sy-datum.
  PARAMETERS: p_data_f TYPE sy-datum.

  SELECTION-SCREEN SKIP.

  SELECTION-SCREEN BEGIN OF BLOCK a4 WITH FRAME TITLE TEXT-002.
    PARAMETERS: p_m AS CHECKBOX DEFAULT 'X'.
  SELECTION-SCREEN END OF BLOCK a4.

SELECTION-SCREEN END OF BLOCK a3.

SELECTION-SCREEN BEGIN OF BLOCK a5 WITH FRAME TITLE TEXT-003.
  PARAMETERS: p_email AS CHECKBOX DEFAULT ''.
  SELECT-OPTIONS s_email FOR adr6-smtp_addr NO INTERVALS.
SELECTION-SCREEN END OF BLOCK a5.

DATA: p_data        TYPE sy-datum.

DATA: v_file(128)   TYPE C,
      v_data        TYPE d,
      v_email_corpo TYPE string,
      taxa          TYPE string,
      v_tx          TYPE p DECIMALS 6.

TYPES: BEGIN OF TEXT,
  LINE(120),
END OF TEXT.

DATA: response_headers(80) OCCURS 0 WITH HEADER LINE.
DATA: response             TYPE TABLE OF TEXT WITH HEADER LINE,
      response_entity_body TYPE TABLE OF TEXT WITH HEADER LINE.
DATA: v_response_entity_body(120) TYPE C.
DATA: v_len   TYPE I,
      btocrlf TYPE C VALUE 'Y'.
DATA: t_text(120) OCCURS 0 WITH HEADER LINE.

TYPES: BEGIN OF y_linhas,
  linha TYPE string,
END OF y_linhas.

TYPES: BEGIN OF y_bcb,
  DATA(10)           TYPE C,
        codmoeda(3)        TYPE C,
        tipo(1)            TYPE C,
        moeda(3)           TYPE C,
        taxacompra(15)     TYPE C,
        taxavenda(15)      TYPE C,
        paridadecompra(15) TYPE C,
        paridadevenda(15)  TYPE C,
END OF y_bcb.

TYPES: BEGIN OF y_curr,
  COUNT TYPE I,
  ukurs TYPE tcurr-ukurs,
END OF y_curr.

DATA: BEGIN OF t_linhas OCCURS 1,
  linha(500) TYPE C,
END OF t_linhas.

DATA: t_exch_rate_list  TYPE STANDARD TABLE OF bapi1093_0 WITH HEADER LINE.

DATA: wa_return TYPE bapiret2,
      wa_curr   TYPE y_curr.

DATA:  t_bcb    TYPE STANDARD TABLE OF y_bcb WITH HEADER LINE.

DATA: v_linha(60000)   TYPE C,
      v_csv(60000)     TYPE C,
      v_xml(60000)     TYPE C,
      v_status_text(2) TYPE C,
      v_usd_eur        TYPE string,
      v_brl_usd        TYPE string,
      v_data_email(10)     TYPE C.

DATA: v_date    TYPE scal-DATE,
      v_facdate TYPE scal-DATE.

TYPES: BEGIN OF y_tcurr.
  INCLUDE STRUCTURE tcurr.
TYPES: END OF   y_tcurr.

DATA: t_tcurr TYPE SORTED TABLE OF y_tcurr WITH UNIQUE KEY kurst fcurr tcurr.

DATA: wa_tcurr TYPE y_tcurr.

DATA : v_begda_ex             TYPE C,
      v_endda_ex             TYPE C,
      v_begda                TYPE d,
      v_endda                TYPE d,
      v_prox_dia             TYPE d,
      v_ultimodiautil        TYPE d,
      v_ultimodiautil_ex(10) TYPE C,
      v_diautil(1)           TYPE C.

**-------------------------------------------------------------------**
** VALIDA PARÂMETROS DE SELEÇÃO
**-------------------------------------------------------------------**
AT SELECTION-SCREEN.

IF p_data_i IS INITIAL.
  p_data_i = sy-datum.
ENDIF.

IF p_data_f IS NOT INITIAL AND
p_data_f LT p_data_i.
  MESSAGE 'Erro, data final menor que data inicial.' TYPE 'E' DISPLAY LIKE 'E'.
ENDIF.

**-------------------------------------------------------------------**
** START-OF-SELECTION.
**-------------------------------------------------------------------**
START-OF-SELECTION.

PERFORM cabecalho_mail.

IF p_data_f IS INITIAL.
  p_data_f = p_data_i.
ENDIF.

p_data = p_data_i.

WHILE p_data LE p_data_f.


  IF p_data IS INITIAL.
    p_data = sy-datum.
  ENDIF.

  v_data = p_data.
  v_date = p_data.

  CALL FUNCTION 'DATE_CONVERT_TO_FACTORYDATE'
  EXPORTING
    correct_option      = '-'
    DATE                = v_date
    factory_calendar_id = 'BR'
  IMPORTING
    DATE                = v_facdate
  EXCEPTIONS
    OTHERS              = 7.

  IF v_facdate EQ v_date. " Verifica se é dia útil

    v_diautil = 'X'.
    PERFORM interface_bcb.

  ELSE.

    CONCATENATE v_facdate+6(2) v_facdate+4(2) v_facdate(4) INTO v_ultimodiautil.
    CALL FUNCTION 'CONVERSION_EXIT_INVDT_INPUT'
    EXPORTING
      INPUT  = v_ultimodiautil
    IMPORTING
      OUTPUT = v_ultimodiautil_ex.

    "Busca os valores de todos as cotações do ultimo dia útil a serem lançados nos fim de samana e feriados
    SELECT * FROM tcurr INTO TABLE t_tcurr WHERE gdatu EQ v_ultimodiautil_ex.

  ENDIF.

  IF  v_status_text EQ 'OK' OR v_diautil NE 'X'.

    "Categoria M
    IF p_m IS NOT INITIAL.
      PERFORM header_cat USING 'M'.

      PERFORM lanca_taxa_bcb   USING 'USD' 'BRL' 'M' 'A' '1' '' 'X'.
      PERFORM lanca_taxa_bcb2  USING 'BRL' 'USD' 'M' 'A' '1' '' ''.

      PERFORM lanca_taxa_bcb   USING 'EUR' 'BRL' 'M' 'B' '1' '' 'X'.
      PERFORM lanca_taxa_bcb2  USING 'BRL' 'EUR' 'M' 'B' '1' '' ''.

      PERFORM lanca_taxa_bcb   USING 'GBP' 'BRL' 'M' 'B' '1' '' 'X'.
      PERFORM lanca_taxa_bcb2  USING 'BRL' 'GBP' 'M' 'B' '1' '' ''.

      PERFORM fecha_cat.
    ENDIF.
  ENDIF.

  p_data = p_data + 1.
ENDWHILE.

PERFORM fecha_cat.
PERFORM mensal.

IF p_email IS NOT INITIAL.

  PERFORM envia_mail.

ENDIF.

*&---------------------------------------------------------------------*
*&      Form  lanca_taxa_bcb - Para moeda x REAL
*&---------------------------------------------------------------------*
FORM lanca_taxa_bcb USING p_moeda_origem p_moeda_destino p_categoria p_cor p_fator_0rigem p_paridade p_venda.

  DATA v_moed TYPE string.

  CLEAR: wa_return, taxa.
  CONCATENATE p_moeda_origem 'X' p_moeda_destino INTO v_moed SEPARATED BY space.

  IF v_diautil EQ 'X'.

    CLEAR: t_bcb.

    IF p_paridade IS INITIAL.

      READ TABLE t_bcb WITH KEY moeda = p_moeda_origem.

      CHECK sy-subrc = 0.

      IF p_venda IS INITIAL.
        REPLACE ',' WITH '.' INTO t_bcb-taxacompra .
        taxa = t_bcb-taxacompra * p_fator_0rigem.
      ELSE.
        REPLACE ',' WITH '.' INTO t_bcb-taxavenda .
        taxa = t_bcb-taxavenda * p_fator_0rigem.
      ENDIF.

    ELSE.

      READ TABLE t_bcb WITH KEY moeda = p_moeda_destino.

      CHECK sy-subrc = 0.

      IF p_venda IS INITIAL.
        REPLACE ',' WITH '.' INTO t_bcb-paridadecompra.
        taxa = t_bcb-paridadecompra.
      ELSE.
        REPLACE ',' WITH '.' INTO t_bcb-paridadevenda.
        taxa = t_bcb-paridadevenda.
      ENDIF.

    ENDIF.

    CALL FUNCTION 'ZFI_COTACAO_MOEDAS'
    EXPORTING
      taxa_cambio   = taxa
      moeda_destino = p_moeda_destino
      moeda_origem  = p_moeda_origem
      categoria     = p_categoria
    DATA          = v_data
          fator_origem  = p_fator_0rigem
          fator_destino = '1'
    IMPORTING
      RETURN        = wa_return.

    PERFORM trata_retorno USING p_moeda_origem p_categoria v_data.

    IF p_cor = 'A'.
      PERFORM taxa_azul USING v_moed taxa.
  ELSEIF p_cor = 'B'.
      PERFORM taxa_branca USING v_moed taxa.
    ENDIF.

  ELSE.

    READ TABLE t_tcurr INTO wa_tcurr WITH KEY kurst = p_categoria fcurr = p_moeda_origem tcurr = p_moeda_destino.

    IF sy-subrc EQ 0.

      taxa = wa_tcurr-ukurs.

      CALL FUNCTION 'ZFI_COTACAO_MOEDAS'
      EXPORTING
        taxa_cambio   = taxa
        moeda_destino = p_moeda_destino
        moeda_origem  = p_moeda_origem
        categoria     = p_categoria
      DATA          = v_data
            fator_origem  = p_fator_0rigem
            fator_destino = '1'
      IMPORTING
        RETURN        = wa_return.

      PERFORM trata_retorno USING p_moeda_origem p_categoria v_data.

    ELSE.

      taxa = 'Taxa do ultimo dia útil não foi cadastrada.'.

    ENDIF.

    IF p_cor = 'A'.
      PERFORM taxa_azul USING v_moed taxa.
  ELSEIF p_cor = 'B'.
      PERFORM taxa_branca USING v_moed taxa.
    ENDIF.

  ENDIF.

ENDFORM.                    "LANCA_TAXA_BCB


*&---------------------------------------------------------------------*
*&      Form  lanca_taxa_bcb2 - para REAL X ...
*&---------------------------------------------------------------------*
FORM lanca_taxa_bcb2 USING p_moeda_origem p_moeda_destino p_categoria p_cor p_fator_0rigem p_paridade p_venda.

  DATA v_moed TYPE string.

  CLEAR: wa_return, taxa.
  CONCATENATE p_moeda_origem 'X' p_moeda_destino INTO v_moed SEPARATED BY space.

  IF v_diautil EQ 'X'.

    CLEAR: t_bcb.

    IF p_paridade IS INITIAL.

      READ TABLE t_bcb WITH KEY moeda = p_moeda_destino.

      CHECK sy-subrc = 0.

      IF p_venda IS INITIAL.
        REPLACE ',' WITH '.' INTO t_bcb-taxacompra .
        taxa = t_bcb-taxacompra * p_fator_0rigem.
      ELSE.
        REPLACE ',' WITH '.' INTO t_bcb-taxavenda .
        taxa = t_bcb-taxavenda * p_fator_0rigem.
      ENDIF.

    ELSE.

      READ TABLE t_bcb WITH KEY moeda = p_moeda_origem.

      CHECK sy-subrc = 0.

      IF p_venda IS INITIAL.
        REPLACE ',' WITH '.' INTO t_bcb-paridadecompra.
        taxa = t_bcb-paridadecompra.
      ELSE.
        REPLACE ',' WITH '.' INTO t_bcb-paridadevenda.
        taxa = t_bcb-paridadevenda.
      ENDIF.

    ENDIF.

    CALL FUNCTION 'ZFI_COTACAO_MOEDAS'
    EXPORTING
      taxa_cambio   = taxa
      moeda_destino = p_moeda_destino
      moeda_origem  = p_moeda_origem
      categoria     = p_categoria
      DATA          = v_data
      fator_origem  = p_fator_0rigem
      fator_destino = '1'
    IMPORTING
      RETURN        = wa_return.

    PERFORM trata_retorno USING p_moeda_origem p_categoria v_data.

    IF p_cor = 'A'.
      PERFORM taxa_azul USING v_moed taxa.
  ELSEIF p_cor = 'B'.
      PERFORM taxa_branca USING v_moed taxa.
    ENDIF.

  ELSE.

    READ TABLE t_tcurr INTO wa_tcurr WITH KEY kurst = p_categoria fcurr = p_moeda_origem tcurr = p_moeda_destino.

    IF sy-subrc EQ 0.

      taxa = wa_tcurr-ukurs.

      CALL FUNCTION 'ZFI_COTACAO_MOEDAS'
      EXPORTING
        taxa_cambio   = taxa
        moeda_destino = p_moeda_destino
        moeda_origem  = p_moeda_origem
        categoria     = p_categoria
      DATA          = v_data
            fator_origem  = p_fator_0rigem
            fator_destino = '1'
      IMPORTING
        RETURN        = wa_return.

      PERFORM trata_retorno USING p_moeda_origem p_categoria v_data.

    ELSE.

      taxa = 'Taxa do ultimo dia útil não foi cadastrada.'.

    ENDIF.

    IF p_cor = 'A'.
      PERFORM taxa_azul USING v_moed taxa.
  ELSEIF p_cor = 'B'.
      PERFORM taxa_branca USING v_moed taxa.
    ENDIF.

  ENDIF.

ENDFORM.                    "LANCA_TAXA_BCB


*&---------------------------------------------------------------------*
*&      Form  envia_mail
*&---------------------------------------------------------------------*
FORM envia_mail.

  DATA: lr_mail_data    TYPE REF TO cl_crm_email_data,
        ls_struc_mail   TYPE crms_email_mime_struc,
        lv_send_request TYPE sysuuid_x,
        lv_to           TYPE crms_email_recipient.

*   Cria a Mensagem de E-mail
  CREATE OBJECT lr_mail_data.

*   Preenche o Remetente.
  lr_mail_data->from-address = 'moedas@goiasa.com.br'. "'Integração de cotação de moedas.
  lr_mail_data->from-name    = lr_mail_data->from-address.
  lr_mail_data->from-ID      = lr_mail_data->from-address.

*   Preenche os Destinatários.
  LOOP AT s_email.
    lv_to-address = s_email-low.
    lv_to-name    = sy-tabix.
    APPEND lv_to TO lr_mail_data->to.
  ENDLOOP.

    v_data = sy-datum.
    v_data_email = |{ v_data+6(2) }/{ v_data+4(2) }/{ v_data(4) }|.

*   Define o Assunto do E-mail
  CONCATENATE 'COTAÇÃO DE MOEDAS -' v_data_email INTO  lr_mail_data->subject SEPARATED BY space.

*   Define o Assunto do E-mail
  IF sy-sysid EQ 'ECQ'.
    CONCATENATE lr_mail_data->subject 'Ambiente Qualidade' INTO lr_mail_data->subject SEPARATED BY space.
ELSEIF sy-sysid EQ 'ECP'.
    CONCATENATE lr_mail_data->subject 'Ambiente Produção' INTO lr_mail_data->subject SEPARATED BY space.
  ENDIF.

*   Mensagem no Corpo do E-mail.
  ls_struc_mail-mime_type     = 'text/html'.
  ls_struc_mail-file_name     = 'body.htm'.
  ls_struc_mail-content_ascii =  v_email_corpo.
  APPEND ls_struc_mail TO lr_mail_data->body.

*   Envia o E-mail
  lv_send_request = cl_crm_email_utility_base=>send_email( iv_mail_data = lr_mail_data ).

* Processa o Envio Imediato
  SUBMIT rsconn01 AND RETURN.

  IF sy-subrc = 0.

    MESSAGE 'Email enviado com sucesso!' TYPE 'S' DISPLAY LIKE 'S'.

  ENDIF.

ENDFORM.                    "envia_mail
*&---------------------------------------------------------------------*
*&      Form  INTERFACE_BCB
*&---------------------------------------------------------------------*
FORM interface_bcb .

  CONCATENATE 'http://www4.bcb.gov.br/Download/fechamento/' v_data+0(4) v_data+4(2) v_data+6(2) '.csv' INTO v_file.

  CALL FUNCTION 'HTTP_GET'
  EXPORTING
    absolute_uri                = v_file
*     REQUEST_ENTITY_BODY_LENGTH  =
*     rfc_destination             = 'SAPHTTP'
*     PROXY                       =
*     PROXY_USER                  =
*     PROXY_PASSWORD              =
*     USER                        =
*     PASSWORD                    =
    blankstocrlf                = btocrlf
*     TIMEOUT                     =
  IMPORTING
*     STATUS_CODE                 =
    status_text                 = v_status_text
    response_entity_body_length = v_len
  TABLES
*     REQUEST_ENTITY_BODY         =
    response_entity_body        = response
    response_headers            = response_headers
*     REQUEST_HEADERS             =
  EXCEPTIONS
    connect_failed              = 1
    timeout                     = 2
    internal_error              = 3
    tcpip_error                 = 4
    data_error                  = 5
    system_failure              = 6
    communication_failure       = 7
    OTHERS                      = 8.
  IF v_status_text NE 'OK'.
    CONCATENATE v_email_corpo '<br />' 'Erro de comunicação com o Banco Central do Brasil' INTO v_email_corpo SEPARATED BY space.
    "Caso o Arquivo não tenha sido encontrado no Banco Central
*  PERFORM LOG USING SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ELSE.

      t_text[]  = response[].

      LOOP AT t_text.
        CONCATENATE v_csv t_text INTO v_csv.
      ENDLOOP.

      LOOP AT t_text INTO v_linha.
        SPLIT v_linha AT ';' INTO: t_bcb-DATA
              t_bcb-codmoeda
              t_bcb-tipo
              t_bcb-moeda
              t_bcb-taxacompra
              t_bcb-taxavenda
              t_bcb-paridadecompra
              t_bcb-paridadevenda.
        APPEND t_bcb.

      ENDLOOP.
      SORT t_bcb BY moeda.
    ENDIF.
  ENDFORM.                    " INTERFACE_BCB

*&---------------------------------------------------------------------*
*&      Form  taxa_azul
*&---------------------------------------------------------------------*
  FORM taxa_azul USING moedas tax.

    CONCATENATE v_email_corpo
    '<tr height="20" style="height: 15.0pt; font-size: 11.0pt; color: black; font-weight: 400;'
    ' text-decoration: none; text-underline-style: none; text-line-through: none; font-family: Calibri;'
    ' border-top: .5pt solid #95B3D7; border-right: none; border-bottom: .5pt solid #95B3D7; border-left: .5pt solid #95B3D7;'
    ' background: #DCE6F1; mso-pattern: #DCE6F1 none">'
    ' <td width="170" style="border-style: none solid solid; border-color: -moz-use-text-color windowtext rgb(149, 179, 215) rgb(149, 179, 215); '
    'border-width: medium 1pt 1pt; padding: 0in 5.4pt; height: 15pt;" nowrap="nowrap" valign="bottom" >'
    '<p class="MsoNormal"><span style="font-size: 11pt; font-family: &quot;Calibri&quot;,&quot;sans-serif&quot;; color: black;" lang="PT-BR">'
    moedas
    '</span></p>'
    '</td>'
    '<td style="border-style: none solid solid none;'
    ' border-color: -moz-use-text-color rgb(149, 179, 215) rgb(149, 179, 215) -moz-use-text-color;'
    ' border-width: medium 1pt 1pt medium; padding: 0in 5.4pt; height: 15pt;" nowrap="nowrap" valign="bottom" >'
    '<p class="MsoNormal" style="text-align: right;" align="right">'
    '<span style="font-size: 11pt; font-family: &quot;Calibri&quot;,&quot;sans-serif&quot;; color: black;" lang="PT-BR">  '
    tax
    '<o:p></o:p></span></p>'
    '</td>'
    '</tr>'
    INTO v_email_corpo SEPARATED BY space.

  ENDFORM.                    "taxa_azul
*&---------------------------------------------------------------------*
*&      Form  taxa_branca
*&---------------------------------------------------------------------*
  FORM taxa_branca USING moedas tax.

    CONCATENATE v_email_corpo
    '<tr height="20" style="height: 15.0pt; font-size: 11.0pt;'
    ' color: black; font-weight: 400; text-decoration: none;'
    ' text-underline-style: none; text-line-through: none; font-family: Calibri;'
    ' border-top: .5pt solid #95B3D7; border-right: none; border-bottom: .5pt solid #95B3D7;'
    ' border-left: .5pt solid #95B3D7;  mso-pattern: #DCE6F1 none">'
    '<td width="170" style="border-style: none solid solid;'
    ' border-color: -moz-use-text-color windowtext rgb(149, 179, 215) rgb(149, 179, 215);'
    ' border-width: medium 1pt 1pt; padding: 0in 5.4pt; '
    ' height: 15pt;" nowrap="nowrap" valign="bottom" >'
    '<p class="MsoNormal"><span style="font-size: 11pt;'
    ' font-family: &quot;Calibri&quot;,&quot;sans-serif&quot;; color: black;" lang="PT-BR">'
    moedas
    '</span></p>'
    '</td>'
    '<td style="border-style: none solid solid none;'
    ' border-color: -moz-use-text-color rgb(149, 179, 215) rgb(149, 179, 215) -moz-use-text-color;'
    ' border-width: medium 1pt 1pt medium; padding: 0in 5.4pt; height: 15pt;" nowrap="nowrap" valign="bottom" >'
    '<p class="MsoNormal" style="text-align: right;" align="right"><span style="font-size: 11pt;'
    ' font-family: &quot;Calibri&quot;,&quot;sans-serif&quot;; color: black;" lang="PT-BR">'
    taxa
    '<o:p></o:p></span></p>'
    '</td>'
    '</tr>'
    INTO v_email_corpo SEPARATED BY space.

  ENDFORM.                    "taxa_branca
*&---------------------------------------------------------------------*
*&      Form  HEADER_CAT
*&---------------------------------------------------------------------*
  FORM header_cat USING categoria.

    CONCATENATE v_email_corpo
    '<table border="1" cellpadding="1" cellspacing="0" >'
    '<tr height="20" style="height:15.0pt;font-size: 11.0pt; color: white; '
    'font-weight: 700; text-decoration: none; text-underline-style: none; text-line-through: none;'
    ' font-family: Calibri; border-top: .5pt solid #95B3D7; border-right: none; border-bottom: .5pt solid #95B3D7;'
    ' border-left: .5pt solid #95B3D7; background: #4F81BD; mso-pattern: #4F81BD none">'
    '<td height="20" width="170" style="border-style: none solid solid none;'
    ' border-color: -moz-use-text-color rgb(149, 179, 215) rgb(149, 179, 215) -moz-use-text-color;'
    ' border-width: medium 1pt 1pt medium; padding: 0in 5.4pt; width: 58pt; height: 15pt;" nowrap="nowrap"'
    ' valign="bottom" width="77">'
    'Categoria'  categoria '</td>'
    '<td style="border-style: none solid solid none;'
    ' border-color: -moz-use-text-color rgb(149, 179, 215) rgb(149, 179, 215) -moz-use-text-color;'
    ' border-width: medium 1pt 1pt medium; padding: 0in 5.4pt; height: 15pt;"'
    ' nowrap="nowrap" valign="bottom" >'
    '    Taxa</td>'
    '</tr>'  INTO v_email_corpo SEPARATED BY space.

  ENDFORM.                    "HEADER_CAT

*&---------------------------------------------------------------------*
*&      Form  MENSAL
*&---------------------------------------------------------------------*
  FORM mensal .

    DATA: v_endda_formatado(10) TYPE C,
          v_begda_formatado(10) TYPE C.

    CLEAR: v_endda, v_endda, v_endda_ex, v_begda_ex, v_prox_dia.

    CONCATENATE v_data(4) v_data+4(2) '01' INTO v_begda.
    CONCATENATE v_data(4) v_data+4(2) v_data+6(2) INTO v_endda.

    v_prox_dia = v_endda + 1.

    IF ( v_prox_dia+6(2) EQ 01 ). "A data do processamento seja o ultimo dia do mes

      v_begda_formatado = |{ v_begda+6(2) }/{ v_begda+4(2) }/{ v_begda(4) }|.

      v_endda_formatado = |{ v_endda+6(2) }/{ v_endda+4(2) }/{ v_endda(4) }|.

      PERFORM mensais_mail.

      IF v_usd_eur IS NOT INITIAL.
        PERFORM header_cat USING 'M'.
        PERFORM fecha_cat.

      ENDIF.

      CALL FUNCTION 'CONVERSION_EXIT_INVDT_INPUT'
      EXPORTING
        INPUT  = v_begda_formatado
      IMPORTING
        OUTPUT = v_begda_ex.

      CALL FUNCTION 'CONVERSION_EXIT_INVDT_INPUT'
      EXPORTING
        INPUT  = v_endda_formatado
      IMPORTING
        OUTPUT = v_endda_ex.

      PERFORM header_cat USING '1002'.

      PERFORM taxa_media  USING 'USD' 'BRL' 'M' '1002' 'A' '1' ''.
      PERFORM taxa_media  USING 'BRL' 'USD' 'M' '1002' 'A' '1' ''.

      PERFORM taxa_media  USING 'EUR' 'BRL' 'M' '1002' 'B' '1' ''.
      PERFORM taxa_media  USING 'BRL' 'EUR' 'M' '1002' 'B' '1' ''.

      PERFORM taxa_media  USING 'GBP' 'BRL' 'M' '1002' 'B' '1' ''.
      PERFORM taxa_media  USING 'BRL' 'GBP' 'M' '1002' 'B' '1' ''.

      PERFORM fecha_cat.

    ENDIF.

  ENDFORM.                    " MENSAL
*&---------------------------------------------------------------------*
*&      Form  TAXA_MEDIA
*&---------------------------------------------------------------------*
  FORM taxa_media  USING p_moeda_origem p_moeda_destino p_categoria_origem p_categoria_destino p_cor p_fator_0rigem p_paridade.

    DATA v_moed TYPE string.

    CLEAR: taxa, wa_return.

    CONCATENATE p_moeda_origem 'X' p_moeda_destino INTO v_moed SEPARATED BY space.

    CLEAR wa_curr.
    SELECT COUNT(*) AVG( ukurs ) UP TO 1 ROWS
    FROM tcurr INTO wa_curr
    WHERE kurst = p_categoria_origem   AND
    fcurr = p_moeda_origem AND
    tcurr = p_moeda_destino AND
    gdatu BETWEEN v_endda_ex AND v_begda_ex.

    IF wa_curr-COUNT EQ v_endda+6(2).

      taxa = wa_curr-ukurs.

      CLEAR wa_return.

      CALL FUNCTION 'ZFI_COTACAO_MOEDAS'
      EXPORTING
        taxa_cambio   = taxa
        moeda_destino = p_moeda_destino
        moeda_origem  = p_moeda_origem
        categoria     = p_categoria_destino
        DATA          = v_data
        fator_origem  = '1'
        fator_destino = '1'
      IMPORTING
        RETURN        = wa_return.

      PERFORM trata_retorno USING p_moeda_origem p_categoria_destino v_data.

      IF p_cor = 'A'.
        PERFORM taxa_azul USING v_moed taxa.
    ELSEIF p_cor = 'B'.
        PERFORM taxa_branca USING v_moed taxa.
      ENDIF.

    ELSE.

      taxa = 'Não foram lançadas todas as taxas para que se faça uma media'.

      IF p_cor = 'A'.
        PERFORM taxa_azul USING v_moed taxa.
    ELSEIF p_cor = 'B'.
        PERFORM taxa_branca USING v_moed taxa.
      ENDIF.

    ENDIF.

  ENDFORM.                    " TAXA_MEDIA
*&---------------------------------------------------------------------*
*&      Form  CABECALHO_MAIL
*&---------------------------------------------------------------------*
  FORM cabecalho_mail.

    v_data = sy-datum.
    v_data_email = |{ v_data+6(2) }/{ v_data+4(2) }/{ v_data(4) }|.

    CONCATENATE v_email_corpo
    '<b>'
    '<span style="color:#1F497D; font-family:Calibri; font-size:larger">Cotações Atualizadas em'
    v_data_email
    '<br />'
    '</span>'
    '</b>'
    '<hr />'
    '<BR />' INTO v_email_corpo SEPARATED BY space.

  ENDFORM.                    "CABECALHO_MAIL
*&---------------------------------------------------------------------*
*&      Form  MENSAIS_MAIL
*&---------------------------------------------------------------------*
  FORM mensais_mail.

    v_data_email = | { v_data+6(2) }/{ v_data+4(2) }/{ v_data(4) } |.

    CONCATENATE v_email_corpo
    '<br />'
    '<b>'
    '<span style="color:#1F497D; font-family:Calibri; font-size:medium">Taxas Mensais'
    '<br />'
    '</span>'
    '</b>'
    '<hr />'
    '<BR />' INTO v_email_corpo SEPARATED BY space.

  ENDFORM.                    "CABECALHO_MAIL

*&---------------------------------------------------------------------*
*&      Form  TRATA_RETORNO
*&---------------------------------------------------------------------*
  FORM trata_retorno USING p_moeda_origem p_categoria v_data.

    DATA: v_wkset  TYPE curr_wkset.

    IF wa_return-TYPE EQ 'E'.
      taxa = wa_return-MESSAGE+57.
    ELSE.

      v_tx = taxa.
      taxa = v_tx.

    IF   p_moeda_origem EQ 'USD' AND p_categoria EQ 'M'.
        v_wkset  = 'USD_BRL_M'.

    ELSEIF p_moeda_origem EQ 'EUR' AND p_categoria EQ 'M'.
        v_wkset  = 'EUR_BRL_M'.

    ELSEIF p_moeda_origem EQ 'GBP' AND p_categoria EQ 'M'.
        v_wkset  = 'GBP_BRL_M'.

      ENDIF.

      IF v_wkset IS NOT INITIAL.

        SELECT * FROM tcurwkdat
        INTO @DATA(ls_tcurwkdat)
              WHERE wkset EQ @v_wkset.
*              AND  wkset EQ @v_iwkset.
        ENDSELECT.

        IF ls_tcurwkdat IS NOT INITIAL
        AND ls_tcurwkdat-lastmaint < v_data.

          ls_tcurwkdat-lastmaint = v_data.
          MODIFY tcurwkdat FROM ls_tcurwkdat.

        ENDIF.

      ENDIF.

    ENDIF.

  ENDFORM.                    " TRATA_RETORNO

*&---------------------------------------------------------------------*
*&      Form  FECHA_CAT
*&---------------------------------------------------------------------*
  FORM fecha_cat.

    CONCATENATE v_email_corpo
    '</table>' '<br />'
    INTO v_email_corpo SEPARATED BY space.

  ENDFORM.                    "FECHA_CAT
