/**
 * buildTokenMap_
 * ----------------------------------------------------
 * Constrói o mapa FINAL de tokens → valores
 * Sistema multi-país escalável
 */
function buildTokenMap_(sheet, row, opts) {

  opts = opts || {};

  const pais        = String(opts.pais || 'BR').toUpperCase();
  const tipoEstacao = String(opts.tipoEstacao || '').toUpperCase();
  const v2Juridico  = !!opts.v2Juridico;

  const { get } = makeRowAccessor_(sheet, row);

  // =====================================================
  // 🌍 I18N ENGINE
  // =====================================================

  const I18N = {

    BR: {
      TITULO: 'Croqui de Implementação de Estação',
      CAPACIDADE: 'Capacidade',
      DIMENSOES: 'Dimensões',
      AREA_TOTAL: 'Área',
      FAIXA_LIVRE: 'Faixa livre',
      FAIXA_MINIMA: 'Faixa mínima',
      CONDICAO: 'Condição',
      AUTORIZANTE: 'Autorizante',
      CARGO: 'Cargo',
      TELEFONE: 'Telefone',
      EMAIL: 'E-mail',
      DOCUMENTO: 'Documento',
      DATA: 'Data',
      TIPO_PUBLICO: 'Estação em Espaço Público',
      TIPO_PRIVADO: 'Estação em Área Privada',
      BASE_LEGAL: 'Documento técnico elaborado conforme levantamento em campo e diretrizes aplicáveis ao uso do espaço urbano.'
    },

    MX: {
      TITULO: 'Croquis de Implementación de Estación',
      CAPACIDAD: 'Capacidad',
      DIMENSOES: 'Dimensiones',
      AREA_TOTAL: 'Área total',
      FAIXA_LIVRE: 'Franja libre',
      FAIXA_MINIMA: 'Franja mínima',
      CONDICAO: 'Condición',
      AUTORIZANTE: 'Autorizante',
      CARGO: 'Cargo',
      TELEFONE: 'Teléfono',
      EMAIL: 'Correo electrónico',
      DOCUMENTO: 'Documento',
      DATA: 'Fecha',
      TIPO_PUBLICO: 'Estación en Espacio Público',
      TIPO_PRIVADO: 'Estación en Área Privada',
      BASE_LEGAL: 'Documento técnico elaborado conforme regulaciones locales aplicables en México.'
    }

  };

  const L = I18N[pais] || I18N.BR;

  function linha(label, valor, sufixo) {
    if (!valor) return '';
    return `${label}: ${valor}${sufixo || ''}`;
  }

  function formatDateSafe(d) {
    if (!d) return '';
    return Utilities.formatDate(
      new Date(d),
      Session.getScriptTimeZone(),
      pais === 'MX' ? 'dd/MM/yyyy' : 'dd/MM/yyyy'
    );
  }

  // =====================================================
  // 📍 IDENTIDADE
  // =====================================================

  const codigo      = get(COL.Codigo);
  const cidade      = get(COL.Cidade);
  const bairro      = get(COL.Bairro);
  const alcaldia    = get('Alcaldia') || '';
  const subpref     = get(COL.Subprefeitura);
  const endereco    = get(COL.Endereco);
  const localizacao = get(COL.Localizacao);
  const tipoPub     = get(COL.TipoPublica);

  const bairroSubpref = [
    bairro || '',
    subpref ? ' / ' + subpref : ''
  ].join('').trim();

  // =====================================================
  // 🏗️ TÉCNICO (VALORES PUROS)
  // =====================================================

  const capacidade = get(COL.Capacidade);
  const dimensoes  = get(COL.Dimensoes);
  const areaTotal  = get(COL.AreaTotal);
  const faixaLivre = get(COL.Largura);
  const faixaMin   = get(COL.FaixaMinima);
  const condicao   = get(COL.Condicao);

  // =====================================================
  // 🔐 PRIVADO
  // =====================================================

  const autorizante = get(COL.NomeAutorizante);
  const cargo       = get(COL.CargoAutorizante);
  const telefone    = get(COL.TelefoneAutorizante);
  const email       = get(COL.EmailAutorizante);
  const documento   = get(COL.DocumentoAutorizacao);
  const dataAut     = get(COL.DataAutorizacao);

  // =====================================================
  // 🧱 MAPA BASE
  // =====================================================

  const map = {

    '{{TITULO_CROQUI}}' : L.TITULO,

    '{{ID_ESTACAO}}'     : codigo || '',
    '{{CODIGO_ESTACAO}}' : codigo || '',
    '{{CIDADE}}'         : cidade || '',
    '{{BAIRRO_SUBPREFEITURA}}': bairroSubpref,
    '{{ENDERECO}}'       : endereco || '',
    '{{LOCALIZACAO}}'    : localizacao || '',

    '{{TIPO_ESTACAO}}' :
      tipoEstacao === 'PUBLICO'
        ? L.TIPO_PUBLICO
        : L.TIPO_PRIVADO,

    // Técnico
    '{{LINHA_CAPACIDADE}}'   : linha(L.CAPACIDADE || L.CAPACIDAD, capacidade),
    '{{LINHA_DIMENSOES}}'    : linha(L.DIMENSOES, dimensoes),
    '{{LINHA_AREA_TOTAL}}'   : linha(L.AREA_TOTAL, areaTotal, ''),
    '{{LINHA_FAIXA_LIVRE}}'  : linha(L.FAIXA_LIVRE, faixaLivre, ' m'),
    '{{LINHA_FAIXA_MINIMA}}' : linha(L.FAIXA_MINIMA, faixaMin, ' m'),
    '{{LINHA_CONDICAO}}'     : linha(L.CONDICAO, condicao),

    '{{BASE_LEGAL}}' : L.BASE_LEGAL
  };

  // =====================================================
  // 🔒 BLOCO PRIVADO
  // =====================================================

  if (tipoEstacao === 'PRIVADO') {

    Object.assign(map, {

      '{{LINHA_AUTORIZANTE}}' :
        linha(L.AUTORIZANTE, autorizante) ||
        `${L.AUTORIZANTE}: información proporcionada por el responsable legal`,

      '{{LINHA_CARGO}}' :
        linha(L.CARGO, cargo),

      '{{LINHA_TELEFONE}}' :
        linha(L.TELEFONE, telefone),

      '{{LINHA_EMAIL}}' :
        linha(L.EMAIL, email),

      '{{LINHA_DOCUMENTO}}' :
        linha(L.DOCUMENTO, documento),

      '{{LINHA_DATA}}' :
        dataAut
          ? `${L.DATA}: ${formatDateSafe(dataAut)}`
          : `${L.DATA}: información proporcionada`
    });
  }

  // =====================================================
  // 🇲🇽 BLOCO EXTRA MX
  // =====================================================

  if (pais === 'MX') {

    const orgPublica =
      cidade && cidade.toUpperCase().includes('CIUDAD DE MEXICO') && alcaldia
        ? 'Alcaldía de ' + alcaldia
        : 'Municipio de ' + cidade;

    Object.assign(map, {

      '{{ORG_PUBLICA}}' : orgPublica,
      '{{MUNICIPIO}}'   : cidade || '',
      '{{COLONIA}}'     : bairro || '',

      '{{TIPO_PUBLICA}}' :
        tipoPub === 'CALCADA' ? 'Banqueta'
        : tipoPub === 'RUA'   ? 'Vía Pública'
        : ''
    });

    if (v2Juridico) {
      map['{{TEXTO_JURIDICO}}'] =
        tipoEstacao === 'PUBLICO'
          ? TEXTO_JURIDICO_MX_PUBLICO_V2
          : TEXTO_JURIDICO_MX_PRIVADO_V2;
    }
  }

  // =====================================================
  // 🇧🇷 BLOCO EXTRA BR (V2)
  // =====================================================

  if (v2Juridico && pais === 'BR') {

    if (tipoEstacao === 'PUBLICO') {
      Object.assign(map, {
        '{{RESPONSAVEL_TECNICO}}' : 'Equipe técnica JET',
        '{{DATA_LEVANTAMENTO}}'  : formatDateSafe(new Date()),
        '{{OBSERVACOES_TECNICAS}}' :
          'Implantação analisada considerando circulação de pedestres, acessibilidade e segurança viária.'
      });
    }

    if (tipoEstacao === 'PRIVADO') {
      Object.assign(map, {
        '{{DECLARACAO_AUTORIZACAO}}' :
          'A implantação ocorre integralmente em área privada, mediante autorização expressa do responsável legal.',
        '{{RESPONSABILIDADE_CIVIL}}' :
          'A responsabilidade civil pelo uso do espaço é atribuída ao autorizante.',
        '{{VALIDADE_AUTORIZACAO}}' :
          'Autorização válida enquanto mantidas as condições descritas neste documento.'
      });
    }
  }

  return map;
}


const TEXTO_JURIDICO_MX_PUBLICO_V2 =
  'Documento técnico informativo elaborado con base en levantamiento de campo y cartografía oficial de referencia.\n\n' +
  'La instalación de la estación se proyecta en espacio público, considerando circulación peatonal, accesibilidad universal, seguridad vial y normativas municipales aplicables.\n\n' +
  'Este documento no constituye autorización administrativa, permiso ni concesión.';

const TEXTO_JURIDICO_MX_PRIVADO_V2 =
  'La instalación de la estación se realiza íntegramente en propiedad privada, mediante autorización expresa del responsable legal.\n\n' +
  'No existe ocupación de vía pública.\n\n' +
  'La responsabilidad civil y administrativa corresponde exclusivamente al autorizante.';
