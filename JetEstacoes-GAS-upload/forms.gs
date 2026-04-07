/*************************************************
 * CONFIGURAÇÕES GERAIS
 *************************************************/
const ABA_FORMS_BR = 'ENTRADA_FORMS';
const ABA_FORMS_MX = 'ENTRADA_FORMS_MX';

const ABA_ESTACOES = 'Estacoes';
const ABA_SUBPREF = 'MAPA_SUBPREF_SP';

const MAPS_API_KEY = 'AIzaSyD-yITmwagjpKPhAlTX1ecAJA7SNYfae5E';
const DUP_RAIO_METROS = 30;

const DRIVE_PASTA_ESTACOES_ID = '1Qq53khNBF-0ej676WGijT9Vt1c99GLLp';


/*************************************************
 * MAPA DE HEADERS — SUPORTE BR / MX
 * NÃO remover chaves PT-BR (canônicas)
 *************************************************/
const FORMS_HEADER_MAP = {

  // ===== TIPO / STATUS =====
  'Tipo da Estação': 'Tipo da Estação',
  'Tipo de la Estación': 'Tipo da Estação',

  'Tipo da Estação Pública': 'Tipo da Estação Pública',
  'Tipo de Estación Pública': 'Tipo da Estação Pública',

  // ===== LOCALIZAÇÃO =====
  'Localização da Estação': 'Localização da Estação',
  'Ubicación de la Estación': 'Localização da Estação',

  // ===== FOTO =====
  'Foto da Estação': 'Foto da Estação',
  'Foto de la Estación': 'Foto da Estação',

  // ===== DIMENSÕES =====
  'Dimensões da Estação': 'Dimensões da Estação',
  'Dimensiones de la Estación': 'Dimensões da Estação',

  'Largura da Faixa Livre (m)': 'Largura da Faixa Livre (m)',
  'Ancho de Franja Libre (m)': 'Largura da Faixa Livre (m)',

  'Capacidade': 'Capacidade',
  'Capacidad': 'Capacidade',

  // ===== PRIVADO =====
  'Nome do Local Privado': 'Nome do Local Privado',
  'Nombre del Local Privado': 'Nome do Local Privado',

  'Nome do Autorizante': 'Nome do Autorizante',
  'Nombre del Autorizante': 'Nome do Autorizante',

  'Cargo do Autorizante': 'Cargo do Autorizante',
  'Cargo del Autorizante': 'Cargo do Autorizante',

  'Telefone Autorizante': 'Telefone Autorizante',
  'Teléfono del Autorizante': 'Telefone Autorizante',

  'E-mail Autorizante': 'E-mail Autorizante',
  'Correo del Autorizante': 'E-mail Autorizante',

  'Documento de Autorização': 'Documento de Autorização',
  'Documento de Autorización': 'Documento de Autorização',

  // ===== OUTROS =====
  'Observações': 'Observações',
  'Observaciones': 'Observações',

  'Nome do Concorrente (se conhecido)': 'Nome do Concorrente (se conhecido)',
  'Nombre del Competidor (si se conoce)': 'Nome do Concorrente (se conhecido)',

  // ===== PAÍS =====
  'País': 'Pais',
  'País de la Estación': 'Pais'
};


/*************************************************
 * ENTRY POINT — TRIGGER DA PLANILHA
 *************************************************/
function processarEntradaForms(e) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = e.range.getSheet();
    const row = e.range.getRow();

    const abaEstacoes = ss.getSheetByName(ABA_ESTACOES);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values  = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

    const data = {};
    headers.forEach((h, i) => {
      const key = FORMS_HEADER_MAP[h] || h;
      data[key] = values[i];
    });


    // ===== PAÍS (BR / MX) =====
    data['Pais'] = normalizar(data['Pais'] || 'BR');

    // --- TIPO / STATUS ---
    const tipo = normalizar(data['Tipo da Estação']);
    const status = (tipo === 'PUBLICA' || tipo === 'PRIVADA') ? 'SOLICITADO' : 'ATIVO';

    // --- TIPO DE IMPLANTAÇÃO (DERIVADO AUTOMATICAMENTE) ---
    let tipoImplantacao = '';

    if (tipo === 'PUBLICA') {
      tipoImplantacao =
        normalizar(data['Tipo da Estação Pública']) === 'CALCADA'
          ? 'Implantación sobre banqueta'
          : 'Implantación sobre vía pública';
    } else if (tipo === 'PRIVADA') {
      tipoImplantacao = 'Implantación en área privada';
    }


    // --- LAT / LNG ---
    const { lat, lng } = extrairLatLng(data['Localização da Estação']);

  // --- GEO ---
  let geo = {};
  if (lat && lng) {

    geo = reverseGeocode_(lat, lng, data['Pais']);

    if (data['Pais'] === 'MX') {

      let alcaldiaFinal = resolverAlcaldiaMX_(ss, geo.alcaldia);

      if (!alcaldiaFinal && geo.bairro) {
        alcaldiaFinal = resolverAlcaldiaPorColonia_(ss, geo.bairro);
      }

      const coloniaFinal = resolverColoniaMX_(ss, geo.bairro, alcaldiaFinal);
      const mun = resolverMunicipioMX_(ss, geo.cidade);

      geo.alcaldia = alcaldiaFinal || 'SIN_DEFINIR_ENUM';
      geo.bairro   = coloniaFinal || geo.bairro;
      geo.cidade   = mun.municipio || geo.cidade;
      geo.estado   = mun.estado || '';
      geo.regiao   = mun.regiao || '';
    }
  }


    // --- SUBPREF SP ---
    const subpref =
      data['Pais'] === 'BR'
        ? identificarSubprefeituraSP(ss, geo)
        : '';


    // --- FOTO ---
    const fotoUrl = data['Foto da Estação'];
    const fotoFinal = organizarFotoEstacao({
      fotoUrl,
      pais: data['Pais'],
      cidade: geo.cidade,
      bairro: geo.bairro,
      subprefeitura: subpref,
      nomeEstacao: geo.enderecoCompleto
    });


    // --- SEQUÊNCIAS ---
    const seqGlobal = gerarSeqGlobal(abaEstacoes);
    const seqPorBairro = gerarSeqPorBairro(abaEstacoes, geo.bairro);

    // --- DUPLICIDADE ---
    const dup = (lat && lng)
      ? verificarDuplicidade(abaEstacoes, lat, lng, DUP_RAIO_METROS)
      : null;

    const dupGrupo = dup ? dup.codigo : '';
    const dupMotivo = dup ? `DISTANCIA_${dup.distancia}M` : '';

    // ===== CÓDIGO DA ESTAÇÃO (ANTES DO OBJETO) =====
    const bairroCod = normalizar(geo.bairro || 'SEM_BAIRRO')
      .replace(/\s+/g, '_');

    const seqBairroFmt = String(seqPorBairro).padStart(3, '0');

    const codigoEstacao = `${bairroCod}-${seqBairroFmt}`;

    // ===== REGISTRO FINAL =====
    const registro = {
      'RowKey': Utilities.getUuid(),
      'CodigoEstacao': codigoEstacao,
      'Nome da Estação': geo.enderecoCompleto || '',
      'Cidade': geo.cidade || '',
      'Bairro': geo.bairro || '',
      'Endereço completo da estação': geo.enderecoCompleto || '',
      'Subprefeitura': subpref,
      'Localização': data['Localização da Estação'] || '',
      'Latitude': lat || '',
      'Longitude': lng || '',
      'TipoEstacao': tipo,
      'StatusEstacao': status,
      'TipoPublica': normalizar(data['Tipo da Estação Pública'] || ''),
      'Dimensões da Estação': data['Dimensões da Estação'] || '',
      'Largura da Faixa Livre (m)': data['Largura da Faixa Livre (m)'] || '',
      'FaixaLivreMinima': '',
      'Capacidade': data['Capacidade'] || '',
      'AreaTotal': '',
      'CondicaoImplantacao': tipoImplantacao,
      'NomeLocalPrivado': data['Nome do Local Privado'] || '',
      'NomeAutorizante': data['Nome do Autorizante'] || '',
      'CargoAutorizante': data['Cargo do Autorizante'] || '',
      'TelefoneAutorizante': data['Telefone Autorizante'] || '',
      'EmailAutorizante': data['E-mail Autorizante'] || '',
      'DocumentoAutorizacao': data['Documento de Autorização'] || '',
      'Foto da Estação': fotoFinal,
      'Observações': data['Observações'] || '',
      'SeqPorBairro': seqPorBairro,
      'SeqGlobal': seqGlobal,
      'DupGrupo': dupGrupo,
      'DupMotivo': dupMotivo,
      'CriadoPor': Session.getEffectiveUser().getEmail(),
      'OrigemDado': 'FORMS',
      'DataCriacao': new Date(),
      'UltimaEdicao': new Date(),
      'PerfilCriador': 'CAMPO',
      'NomeConcorrente': data['Nome do Concorrente (se conhecido)'] || '',
      'Pais': data['Pais'] === 'MX' ? 'MX' : 'BR',
      'Alcaldia': geo.alcaldia || ''
    };

    escreverPorHeader(abaEstacoes, registro);

  } catch (err) {
    Logger.log(err);
  }
}

/*************************************************
 * UTIL — NORMALIZAÇÃO
 *************************************************/
function normalizar(txt) {
  return (txt || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

/*************************************************
 * UTIL — EXTRAIR LAT / LNG
 *************************************************/
function extrairLatLng(texto) {
  if (!texto) return {};

  // Link Google Maps
  let m = texto.match(/q=(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
  if (m) return { lat: +m[1], lng: +m[2] };

  // Texto "lat,lng"
  m = texto.match(/(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/);
  if (m) return { lat: +m[1], lng: +m[2] };

  return {};
}

/*************************************************
 * GOOGLE MAPS — REVERSE GEOCODE
 *************************************************/
function reverseGeocodeMX_(lat, lng) {

  const url =
    `https://maps.googleapis.com/maps/api/geocode/json?latlng=${lat},${lng}&key=${MAPS_API_KEY}`;

  const res = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  if (!res.results || !res.results.length) return {};

  const comps = res.results[0].address_components;

  const get = type =>
    (comps.find(c => c.types.includes(type)) || {}).long_name || '';

  const estado = get('administrative_area_level_1') || '';

  const cidade =
    get('administrative_area_level_1') ||
    get('locality') ||
    '';

  const alcaldia =
    get('administrative_area_level_3') ||
    get('sublocality_level_1') ||
    get('neighborhood') ||
    '';

  const colonia =
    get('sublocality_level_2') ||
    get('sublocality_level_1') ||
    get('neighborhood') ||
    '';

  return {
    enderecoCompleto: res.results[0].formatted_address,
    cidade: cidade,
    estado: estado,
    alcaldia: alcaldia,
    bairro: colonia
  };
}

function resolverAlcaldiaPorColonia_(ss, colonia) {
  if (!colonia) return '';

  const aba = ss.getSheetByName('ENUM_COLONIAS_CDMX');
  if (!aba) return '';

  const coloniaNorm = normalizar(colonia);
  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const ativo = dados[i][4] === true;
    const normCol = normalizar(dados[i][1]);

    if (ativo && normCol === coloniaNorm) {
      return dados[i][2]; // Alcaldia
    }
  }
  return '';
}

function resolverAlcaldiaMX_(ss, alcaldiaBruta) {

  if (!alcaldiaBruta) return '';

  const aba = ss.getSheetByName('ENUM_ALCALDIAS_MX');
  if (!aba) return '';

  const alvo = normalizar(alcaldiaBruta);
  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const ativo = dados[i][3] === true;
    const norm  = normalizar(dados[i][1]);

    if (ativo && norm === alvo) {
      return dados[i][1]; // nome oficial da alcaldía
    }
  }

  return '';
}


function resolverColoniaMX_(ss, coloniaBruta, alcaldiaOficial) {

  if (!coloniaBruta) return '';

  const aba = ss.getSheetByName('ENUM_COLONIAS_MX');
  if (!aba) return '';

  const alvo = normalizar(coloniaBruta);
  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    const ativo = dados[i][8] === true;
    if (!ativo) continue;

    // força match dentro da alcaldía correta
    if (alcaldiaOficial &&
        normalizar(dados[i][6]) !== normalizar(alcaldiaOficial)) {
      continue;
    }

    const variantes = [
      dados[i][1],
      dados[i][2],
      dados[i][3],
      dados[i][4],
      dados[i][5]
    ].map(v => normalizar(v)).filter(Boolean);

    if (variantes.includes(alvo)) {
      return dados[i][0]; // Colonia oficial
    }
  }

  return '';
}

function resolverMunicipioMX_(ss, cidadeBruta) {

  if (!cidadeBruta) return '';

  const aba = ss.getSheetByName('ENUM_MUNICIPIOS_MX');
  if (!aba) return '';

  const alvo = normalizar(cidadeBruta);
  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    if (dados[i][9] !== true) continue;

    const variantes = [
      dados[i][1],
      dados[i][2],
      dados[i][3],
      dados[i][4],
      dados[i][5]
    ].map(v => normalizar(v)).filter(Boolean);

    if (variantes.includes(alvo)) {
      return {
        municipio: dados[i][0],
        estado: dados[i][6],
        regiao: dados[i][8]
      };
    }
  }

  return {};
}


/*************************************************
 * SUBPREFEITURA — SÃO PAULO
 *************************************************/
function identificarSubprefeituraSP(ss, geo) {
  if (!geo || geo.cidade !== 'São Paulo') return '';

  const bairroNorm = normalizar(geo.bairro);
  if (!bairroNorm) return 'VALIDAR';

  const aba = ss.getSheetByName(ABA_SUBPREF);
  if (!aba) return 'VALIDAR';

  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (normalizar(dados[i][0]) === bairroNorm) {
      return dados[i][1];
    }
  }
  return 'VALIDAR';
}

/*************************************************
 * SEQUÊNCIAS
 *************************************************/
function gerarSeqGlobal(aba) {
  const dados = aba.getDataRange().getValues();
  const header = dados[0];
  const idx = header.indexOf('SeqGlobal');

  let max = 0;
  for (let i = 1; i < dados.length; i++) {
    if (typeof dados[i][idx] === 'number') {
      max = Math.max(max, dados[i][idx]);
    }
  }
  return max + 1;
}

function gerarSeqPorBairro(aba, bairro) {
  if (!bairro) return '';

  const dados = aba.getDataRange().getValues();
  const header = dados[0];
  const iBairro = header.indexOf('Bairro');
  const iSeq = header.indexOf('SeqPorBairro');

  let max = 0;
  for (let i = 1; i < dados.length; i++) {
    if (
      dados[i][iBairro] === bairro &&
      typeof dados[i][iSeq] === 'number'
    ) {
      max = Math.max(max, dados[i][iSeq]);
    }
  }
  return max + 1;
}

/*************************************************
 * DUPLICIDADE — DISTÂNCIA
 *************************************************/
function distanciaMetros(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;

  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);

  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) *
    Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) ** 2;

  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function verificarDuplicidade(aba, lat, lng, raio) {
  const dados = aba.getDataRange().getValues();
  const header = dados[0];

  const iLat = header.indexOf('Latitude');
  const iLng = header.indexOf('Longitude');
  const iCod = header.indexOf('CodigoEstacao');

  for (let i = 1; i < dados.length; i++) {
    const lat2 = dados[i][iLat];
    const lng2 = dados[i][iLng];
    if (!lat2 || !lng2) continue;

    const d = distanciaMetros(lat, lng, lat2, lng2);
    if (d <= raio) {
      return {
        codigo: dados[i][iCod],
        distancia: Math.round(d)
      };
    }
  }
  return null;
}

/*************************************************
 * DRIVE — ORGANIZAÇÃO DE FOTOS
 *************************************************/
function getOuCriarPasta(pastaPai, nome) {
  const it = pastaPai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pastaPai.createFolder(nome);
}

function extrairFileId(url) {
  if (!url) return null;
  const m = url.match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

// mantém o nome EXATO da estação (remove só / e \)
function nomeArquivoSeguro(nome) {
  if (!nome) return 'ESTACAO';
  return nome.replace(/[\/\\]/g, '-').trim();
}

function organizarFotoEstacao({
  fotoUrl,
  pais,
  cidade,
  bairro,
  subprefeitura,
  nomeEstacao
}) {
  if (!fotoUrl) return '';

  const fileId = extrairFileId(fotoUrl);
  if (!fileId) return fotoUrl;

  const file = DriveApp.getFileById(fileId);
  const pastaRaiz = DriveApp.getFolderById(DRIVE_PASTA_ESTACOES_ID);

  // ================= PAÍS =================
  const pastaPais = getOuCriarPasta(
    pastaRaiz,
    normalizar(pais || 'BR')
  );

  // ================= CIDADE =================
  const pastaCidade = getOuCriarPasta(
    pastaPais,
    normalizar(cidade || 'SEM_CIDADE')
  );

  let pastaFinal = pastaCidade;

  // ================= BR =================
  if (
    pais === 'BR' &&
    normalizar(cidade) === 'SAO PAULO' &&
    subprefeitura &&
    subprefeitura !== 'VALIDAR'
  ) {
    pastaFinal = getOuCriarPasta(
      pastaCidade,
      normalizar(subprefeitura)
    );
  }

  // ================= MX =================
  if (pais === 'MX' && bairro) {
    pastaFinal = getOuCriarPasta(
      pastaCidade,
      normalizar(bairro)
    );
  }

  // ================= ARQUIVO =================
  file.setName(nomeArquivoSeguro(nomeEstacao) + '.jpg');
  pastaFinal.addFile(file);

  // remove de outras pastas
  const paisAntigos = file.getParents();
  while (paisAntigos.hasNext()) {
    const p = paisAntigos.next();
    if (p.getId() !== pastaFinal.getId()) {
      p.removeFile(file);
    }
  }

  return file.getUrl();
}


/*************************************************
 * PLANILHA — ESCREVER POR HEADER
 *************************************************/
function escreverPorHeader(aba, registro) {
  const header = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const linha = header.map(col => registro[col] ?? '');
  aba.appendRow(linha);
}


/*************************************************
 * GEO — RESOLVER BR / MX (GLOBAL)
 *************************************************/
function reverseGeocode_(lat, lng, pais) {

  pais = String(pais || '').toUpperCase();

  // 🇲🇽 MÉXICO
  if (pais === 'MX') {
    return reverseGeocodeMX_(lat, lng);
  }

  // 🇧🇷 BRASIL
  return reverseGeocodeBR_(lat, lng);
}


function reverseGeocodeBR_(lat, lng) {

  const url =
    `https://maps.googleapis.com/maps/api/geocode/json?latlng=${lat},${lng}&key=${MAPS_API_KEY}`;

  const res = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  if (!res.results || !res.results.length) return {};

  const comps = res.results[0].address_components;

  const get = type =>
    (comps.find(c => c.types.includes(type)) || {}).long_name || '';

  return {
    enderecoCompleto: res.results[0].formatted_address,
    cidade: get('locality') || get('administrative_area_level_2') || '',
    estado: get('administrative_area_level_1') || '',
    bairro:
      get('sublocality_level_1') ||
      get('neighborhood') ||
      ''
  };
}
