import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

// Armazena na memória RAM para evitar downloads repetitivos
const templateCache = {};

/**
 * Baixa o template ou pega do cache em memória
 */
async function getTemplateBuffer(arquivoTemplate) {
  if (templateCache[arquivoTemplate]) {
    // Retorna uma cópia do buffer para não corromper o cache nas edições
    return templateCache[arquivoTemplate].slice(0);
  }

  const response = await fetch(`/${arquivoTemplate}?v=${new Date().getTime()}`);
  if (!response.ok) {
    throw new Error(`Não foi possível carregar o modelo '${arquivoTemplate}'.`);
  }

  const blob = await response.blob();
  const arrayBuffer = await blob.arrayBuffer();
  
  templateCache[arquivoTemplate] = arrayBuffer;
  return arrayBuffer.slice(0);
}

/**
 * Processa um único documento docx, injetando os dados e baixando para o usuário.
 */
async function processSingleDoc(arquivoTemplate, dadosFormulario, sufixoNome, acessoriosSelecionados = []) {
  const arrayBuffer = await getTemplateBuffer(arquivoTemplate);

  const zip = new PizZip(arrayBuffer);

  // Manipulação Cirúrgica e Segura do XML usando DOMParser
  let xmlStr = zip.file("word/document.xml").asText();
  xmlStr = xmlStr.replace('{ACESSORIOS}', '');

  // O troca.docx usa o formato { CAMPO } com espaços. Normaliza para {{CAMPO}} que o Docxtemplater entende.
  if (arquivoTemplate === 'troca.docx') {
    // Remove espaços internos e converte chave simples em dupla: { MODELO _NOVO } -> {{MODELO_NOVO}}
    xmlStr = xmlStr.replace(/\{\s*([A-Z_]+(?:\s+[A-Z_]+)*)\s*\}/g, (match, p1) => {
      const normalized = p1.replace(/\s+/g, '_').toUpperCase();
      return `{{${normalized}}}`;
    });
  }


  if (arquivoTemplate === 'entrega.docx' && acessoriosSelecionados.length > 0) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, "application/xml");
    
    // Busca todas as tags de texto (<w:t>)
    const textNodes = xmlDoc.getElementsByTagName("w:t");
    let targetNode = null;
    
    for (let i = 0; i < textNodes.length; i++) {
      if (textNodes[i].textContent === 'Fonte') {
        targetNode = textNodes[i];
        break;
      }
    }

    if (targetNode) {
      // Subir na hierarquia até achar o parágrafo pai (<w:p>)
      let pNode = targetNode.parentNode;
      while (pNode && pNode.nodeName !== 'w:p') {
        pNode = pNode.parentNode;
      }

      if (pNode && pNode.parentNode) {
        const parentContainer = pNode.parentNode;
        
        acessoriosSelecionados.forEach(acc => {
           // Monta o XML do novo bullet point (garantindo os namespaces necessários no parsing isolado)
           const bulletTemplateStr = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Standarduser"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="15"/></w:numPr><w:rPr><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>${acc}</w:t></w:r></w:p>`;
           
           const bulletDoc = parser.parseFromString(bulletTemplateStr, "application/xml");
           const importedNode = xmlDoc.importNode(bulletDoc.documentElement, true);
           
           // Insere o novo bullet point cirurgicamente logo APÓS o parágrafo da 'Fonte'
           parentContainer.insertBefore(importedNode, pNode.nextSibling);
           
           // Avança o cursor para injetar o próximo logo abaixo deste
           pNode = importedNode;
        });
      }
    }
    
    const serializer = new XMLSerializer();
    let serialized = serializer.serializeToString(xmlDoc);
    
    // Garante que a declaração XML não seja perdida no parser
    if (!serialized.startsWith("<?xml")) {
      serialized = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serialized;
    }
    xmlStr = serialized;
  }
  
  zip.file("word/document.xml", xmlStr);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true
  });

  doc.render(dadosFormulario);

  const out = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const outputName = `${dadosFormulario.MATRICULA}_${dadosFormulario.NOME}_${dadosFormulario.NUMERO_SERIE}.docx`;
  saveAs(out, outputName);
}

/**
 * Orquestrador principal que lida com a geração de um ou múltiplos documentos
 * baseados no tipo de operação (Entrega, Devolução ou Troca).
 */
export async function processTermos(tipo, isTroca, formDataBase, formDataNovaEntrega, acessoriosSelecionados) {
  if (tipo === 'Entrega') {
    await processSingleDoc('entrega.docx', formDataBase, 'Termo de Entrega', acessoriosSelecionados);
  } else {
    if (isTroca && formDataNovaEntrega) {
      // Troca marcada: gera apenas o Termo de Troca exclusivo com os dados combinados
      const dadosTroca = {
        ...formDataBase,
        TECNICO: formDataBase.NOME_TECNICO, // Alias para o campo {TECNICO} do troca.docx
        // Dados do novo equipamento (disponíveis no template troca.docx)
        MODELO_NOVO: formDataNovaEntrega.MODELO,
        CODIGO_INTERNO_NOVO: formDataNovaEntrega.CODIGO_INTERNO,
        NUMERO_SERIE_NOVO: formDataNovaEntrega.NUMERO_SERIE,
        PATRIMONIO_NOVO: formDataNovaEntrega.PATRIMONIO,
      };
      await processSingleDoc('troca.docx', dadosTroca, 'Termo de Troca');
    } else {
      // Devolução simples: gera apenas o Termo de Devolução normal
      await processSingleDoc('devolucao.docx', formDataBase, 'Termo de Devolução');
    }
  }
}
