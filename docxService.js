import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

const templateCache = {};

async function getTemplateBuffer(arquivoTemplate) {
  if (templateCache[arquivoTemplate]) {
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

async function processSingleDoc(arquivoTemplate, dadosFormulario, sufixoNome, acessoriosSelecionados = []) {
  const arrayBuffer = await getTemplateBuffer(arquivoTemplate);

  const zip = new PizZip(arrayBuffer);

  let xmlStr = zip.file("word/document.xml").asText();
  xmlStr = xmlStr.replace('{ACESSORIOS}', '');

  if (arquivoTemplate === 'entrega.docx' && acessoriosSelecionados.length > 0) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, "application/xml");

    const textNodes = xmlDoc.getElementsByTagName("w:t");
    let targetNode = null;

    for (let i = 0; i < textNodes.length; i++) {
      if (textNodes[i].textContent === 'Fonte') {
        targetNode = textNodes[i];
        break;
      }
    }

    if (targetNode) {
      let pNode = targetNode.parentNode;
      while (pNode && pNode.nodeName !== 'w:p') {
        pNode = pNode.parentNode;
      }

      if (pNode && pNode.parentNode) {
        const parentContainer = pNode.parentNode;

        acessoriosSelecionados.forEach(acc => {
           const bulletTemplateStr = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Standarduser"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="15"/></w:numPr><w:rPr><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>${acc}</w:t></w:r></w:p>`;

           const bulletDoc = parser.parseFromString(bulletTemplateStr, "application/xml");
           const importedNode = xmlDoc.importNode(bulletDoc.documentElement, true);

           parentContainer.insertBefore(importedNode, pNode.nextSibling);

           pNode = importedNode;
        });
      }
    }

    const serializer = new XMLSerializer();
    let serialized = serializer.serializeToString(xmlDoc);

    if (!serialized.startsWith("<?xml")) {
      serialized = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serialized;
    }
    xmlStr = serialized;
  }

  zip.file("word/document.xml", xmlStr);

  const doc = new Docxtemplater(zip, {
    delimiters: { start: '{', end: '}' },
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

export async function processTermos(tipo, isTroca, formDataBase, formDataNovaEntrega, acessoriosSelecionados) {
  if (tipo === 'Entrega') {
    await processSingleDoc('entrega.docx', formDataBase, 'Termo de Entrega', acessoriosSelecionados);
  } else {
    if (isTroca && formDataNovaEntrega) {
      const dadosTroca = {
        ...formDataBase,
        TECNICO: formDataBase.NOME_TECNICO,
        MODELO_NOVO: formDataNovaEntrega.MODELO,
        CODIGO_INTERNO_NOVO: formDataNovaEntrega.CODIGO_INTERNO,
        NUMERO_SERIE_NOVO: formDataNovaEntrega.NUMERO_SERIE,
        PATRIMONIO_NOVO: formDataNovaEntrega.PATRIMONIO,
      };
      await processSingleDoc('troca.docx', dadosTroca, 'Termo de Troca');
    } else {
      await processSingleDoc('devolucao.docx', formDataBase, 'Termo de Devolução');
    }
  }
}
