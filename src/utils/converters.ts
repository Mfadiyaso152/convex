import { jsPDF } from 'jspdf';
import JSZip from 'jszip';
import * as jsYaml from 'js-yaml';
import * as XLSX from 'xlsx';

/**
 * Dynamically loads PDF.js from a fast, reliable CDN
 */
function loadPdfJS(): Promise<any> {
  return new Promise((resolve, reject) => {
    if ((window as any).pdfjsLib) {
      resolve((window as any).pdfjsLib);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js';
    script.onload = () => {
      const pdfjsLib = (window as any).pdfjsLib;
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
      resolve(pdfjsLib);
    };
    script.onerror = () => reject(new Error('فشل في تحميل مكتبة معالجة الـ PDF من الخادم السحابي.'));
    document.head.appendChild(script);
  });
}

/**
 * Escapes unsafe XML characters
 */
function escapeXml(unsafe: string): string {
  return unsafe.replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
      default: return c;
    }
  });
}

/**
 * Generates a clean, valid .docx file structure using JSZip
 */
export async function generateDocx(text: string): Promise<Blob> {
  const zip = new JSZip();

  // 1. [Content_Types].xml
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

  // 2. _rels/.rels
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  // 3. word/document.xml
  const paragraphsXml = text.split('\n').map(line => {
    const cleanLine = escapeXml(line.trim());
    return `<w:p><w:r><w:t>${cleanLine}</w:t></w:r></w:p>`;
  }).join('\n');

  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraphsXml}
  </w:body>
</w:document>`);

  return await zip.generateAsync({ type: 'blob' });
}

/**
 * Parses unstructured text into tabular rows/cells for Excel sheets
 */
export function parseTextToRows(text: string): string[][] {
  const lines = text.split('\n');
  const rows: string[][] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('--- الصفحة')) continue;

    // Split line by multiple spaces, tab, or common punctuation delimiters
    let cells: string[] = [];
    if (trimmed.includes('\t')) {
      cells = trimmed.split('\t');
    } else if (/\s{2,}/.test(trimmed)) {
      cells = trimmed.split(/\s{2,}/);
    } else if (trimmed.includes(';')) {
      cells = trimmed.split(';');
    } else {
      cells = [trimmed];
    }

    const cleanCells = cells.map(c => c.trim()).filter(Boolean);
    if (cleanCells.length > 0) {
      rows.push(cleanCells);
    }
  }
  return rows;
}


/**
 * Extracts raw textual content from PDF pages
 */
export async function convertPdfToText(file: File): Promise<string> {
  const pdfjsLib = await loadPdfJS();
  const arrayBuffer = await readFileAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageText = content.items.map((item: any) => item.str).join(' ');
    text += `--- الصفحة ${i} ---\n${pageText}\n\n`;
  }
  return text.trim();
}

/**
 * Converts first page of a PDF file into a crisp image (PNG/JPG)
 */
export async function convertPdfToImage(
  file: File,
  targetFormat: string
): Promise<{ blob: Blob; url: string; name: string }> {
  const pdfjsLib = await loadPdfJS();
  const arrayBuffer = await readFileAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const page = await pdf.getPage(1); // Convert page 1
  const viewport = page.getViewport({ scale: 2.0 }); // 2x scale for high resolution
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Failed to get 2D canvas context');
  }
  await page.render({ canvasContext: ctx, viewport }).promise;
  const mimeType = targetFormat === 'png' ? 'image/png' : 'image/jpeg';
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, mimeType, 0.95));
  if (!blob) {
    throw new Error('فشل تحويل صفحة المستند إلى صورة');
  }
  const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
  return {
    blob,
    url: URL.createObjectURL(blob),
    name: `${originalBaseName}.${targetFormat}`
  };
}

/**
 * Extracts paragraphs from a standard Word DOCX file using JSZip
 */
export async function convertDocxToText(file: File): Promise<string> {
  const zip = new JSZip();
  const loadedZip = await zip.loadAsync(file);
  const docXmlFile = loadedZip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('الملف ليس ملف Word (DOCX) صالح أو غير مدعوم.');
  }
  const xmlText = await docXmlFile.async('text');
  
  // Parse paragraph nodes from Word XML structure
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, 'text/xml');
  const paragraphs = doc.getElementsByTagName('w:p');
  const textLines: string[] = [];
  
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const textRuns = p.getElementsByTagName('w:t');
    let pText = '';
    for (let j = 0; j < textRuns.length; j++) {
      pText += textRuns[j].textContent || '';
    }
    textLines.push(pText);
  }
  
  return textLines.join('\n').trim();
}

/**
 * Converts Excel sheets (XLSX/XLS) into clean CSV, JSON, HTML, or TXT
 */
export async function convertExcel(
  file: File,
  targetExtension: string
): Promise<{ blob: Blob; url: string; name: string }> {
  const arrayBuffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
  
  if (targetExtension === 'csv') {
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
    return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.csv` };
  }
  
  if (targetExtension === 'json') {
    const jsonContent = XLSX.utils.sheet_to_json(worksheet);
    const jsonStr = JSON.stringify(jsonContent, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
    return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.json` };
  }
  
  if (targetExtension === 'html') {
    const htmlContent = XLSX.utils.sheet_to_html(worksheet);
    const fullHtml = `
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>${file.name}</title>
  <style>
    body { font-family: system-ui, sans-serif; padding: 2rem; background: #f8fafc; color: #1e293b; }
    table { border-collapse: collapse; width: 100%; margin-top: 1rem; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05); }
    th, td { border: 1px solid #e2e8f0; padding: 12px 16px; text-align: right; }
    th { background: #f1f5f9; font-weight: bold; }
    tr:nth-child(even) { background: #f8fafc; }
  </style>
</head>
<body>
  <h2>${originalBaseName}</h2>
  ${htmlContent}
</body>
</html>`;
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.html` };
  }
  
  if (targetExtension === 'txt') {
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
    const tabContent = csvContent.replace(/,/g, '\t'); // Convert commas to tab spaces
    const blob = new Blob([tabContent], { type: 'text/plain;charset=utf-8' });
    return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
  }
  
  throw new Error('صيغة التحويل غير مدعومة لملفات Excel');
}

/**
 * Reads a File as text
 */
export function readFileAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target?.result as string || '');
    reader.onerror = (err) => reject(err);
    reader.readAsText(file);
  });
}

/**
 * Reads a File as ArrayBuffer (useful for zip/images)
 */
export function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target?.result as ArrayBuffer);
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Converts any image File to another image format using HTML Canvas
 */
export function convertImage(
  file: File,
  targetFormat: string
): Promise<{ blob: Blob; url: string; name: string }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth || img.width;
        canvas.height = img.naturalHeight || img.height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          reject(new Error('Failed to get 2D canvas context'));
          return;
        }

        // Fill background with white for JPEG format to avoid black background on transparency
        if (targetFormat === 'jpg' || targetFormat === 'jpeg') {
          ctx.fillStyle = '#FFFFFF';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
        }

        ctx.drawImage(img, 0, 0);

        // Map format to correct MIME type
        let mimeType = 'image/png';
        if (targetFormat === 'jpg' || targetFormat === 'jpeg') {
          mimeType = 'image/jpeg';
        } else if (targetFormat === 'webp') {
          mimeType = 'image/webp';
        } else if (targetFormat === 'bmp') {
          mimeType = 'image/bmp';
        } else if (targetFormat === 'ico') {
          // Canvas doesn't natively support image/x-icon perfectly in all browsers,
          // so we can use image/png as a fallback container or create a smaller canvas
          mimeType = 'image/png';
        }

        if (targetFormat === 'pdf') {
          // Special case: convert image to PDF
          try {
            const pdf = new jsPDF({
              orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
              unit: 'px',
              format: [canvas.width, canvas.height]
            });
            const imgData = canvas.toDataURL('image/jpeg', 0.95);
            pdf.addImage(imgData, 'JPEG', 0, 0, canvas.width, canvas.height);
            const pdfBlob = pdf.output('blob');
            const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
            const newName = `${originalBaseName}.pdf`;
            resolve({
              blob: pdfBlob,
              url: URL.createObjectURL(pdfBlob),
              name: newName
            });
          } catch (err) {
            reject(err);
          }
          return;
        }

        canvas.toBlob((blob) => {
          if (!blob) {
            reject(new Error('Failed to convert image to Blob'));
            return;
          }
          const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
          const newName = `${originalBaseName}.${targetFormat}`;
          resolve({
            blob,
            url: URL.createObjectURL(blob),
            name: newName
          });
        }, mimeType, 0.92);
      };
      img.onerror = (err) => reject(new Error('Failed to load image file'));
      img.src = e.target?.result as string;
    };
    reader.onerror = (err) => reject(err);
    reader.readAsDataURL(file);
  });
}

/**
 * Text to PDF Converter with canvas-based elegant Arabic RTL rendering
 */
export function convertTextToPDF(
  text: string,
  fileName: string
): Promise<{ blob: Blob; url: string; name: string }> {
  return new Promise((resolve, reject) => {
    try {
      const originalBaseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
      const newName = `${originalBaseName}.pdf`;

      // Split text into paragraphs/lines
      const lines = text.split('\n');

      // Create a hidden canvas to measure and draw text beautifully with system RTL support
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        throw new Error('Canvas not supported');
      }

      // PDF dimensions (Standard A4 is roughly 595 x 842 points/pixels at 72dpi, let's use a crisp 1200 x 1700 canvas)
      const canvasWidth = 1200;
      const margin = 80;
      const printableWidth = canvasWidth - margin * 2;
      const fontSize = 24;
      const lineHeight = fontSize * 1.6;

      ctx.font = `${fontSize}px "Tajawal", "Arial", sans-serif`;
      
      // Determine if text is mostly Arabic for RTL
      const arabicPattern = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
      const isArabic = arabicPattern.test(text.substring(0, 500));

      // Wrap text helper
      const wrapText = (txt: string): string[] => {
        const words = txt.split(' ');
        const wrapped: string[] = [];
        let currentLine = '';

        for (let i = 0; i < words.length; i++) {
          const word = words[i];
          const testLine = currentLine ? `${currentLine} ${word}` : word;
          const metrics = ctx.measureText(testLine);
          if (metrics.width > printableWidth && currentLine) {
            wrapped.push(currentLine);
            currentLine = word;
          } else {
            currentLine = testLine;
          }
        }
        if (currentLine) {
          wrapped.push(currentLine);
        }
        return wrapped;
      };

      // Process and wrap all lines
      const allWrappedLines: string[] = [];
      for (const rawLine of lines) {
        if (rawLine.trim() === '') {
          allWrappedLines.push('');
        } else {
          allWrappedLines.push(...wrapText(rawLine));
        }
      }

      // Calculate total height needed. If it exceeds single page height, we will split into multiple canvases (pages).
      const pageHeight = 1700;
      const maxLinesPerPage = Math.floor((pageHeight - margin * 2) / lineHeight);
      
      const pagesCount = Math.max(1, Math.ceil(allWrappedLines.length / maxLinesPerPage));
      const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'px',
        format: [canvasWidth, pageHeight]
      });

      for (let p = 0; p < pagesCount; p++) {
        if (p > 0) {
          pdf.addPage([canvasWidth, pageHeight], 'portrait');
        }

        const pageCanvas = document.createElement('canvas');
        pageCanvas.width = canvasWidth;
        pageCanvas.height = pageHeight;
        const pageCtx = pageCanvas.getContext('2d');
        if (!pageCtx) continue;

        // Fill background
        pageCtx.fillStyle = '#FFFFFF';
        pageCtx.fillRect(0, 0, canvasWidth, pageHeight);

        // Draw header border or clean layout
        pageCtx.strokeStyle = '#E2E8F0';
        pageCtx.lineWidth = 2;
        pageCtx.strokeRect(40, 40, canvasWidth - 80, pageHeight - 80);

        // Setup text styles
        pageCtx.fillStyle = '#1E293B';
        pageCtx.font = `${fontSize}px "Tajawal", "Arial", sans-serif`;
        pageCtx.textBaseline = 'top';

        if (isArabic) {
          pageCtx.textAlign = 'right';
        } else {
          pageCtx.textAlign = 'left';
        }

        const startIdx = p * maxLinesPerPage;
        const endIdx = Math.min(startIdx + maxLinesPerPage, allWrappedLines.length);

        let currentY = margin;

        for (let i = startIdx; i < endIdx; i++) {
          const lineText = allWrappedLines[i];
          if (lineText.trim() !== '') {
            const xPos = isArabic ? canvasWidth - margin : margin;
            pageCtx.fillText(lineText, xPos, currentY);
          }
          currentY += lineHeight;
        }

        // Draw page number in footer
        pageCtx.font = `18px "Tajawal", "Arial", sans-serif`;
        pageCtx.fillStyle = '#94A3B8';
        pageCtx.textAlign = 'center';
        pageCtx.fillText(`صفحة ${p + 1} من ${pagesCount}`, canvasWidth / 2, pageHeight - margin + 20);

        // Add page image to PDF
        const imgData = pageCanvas.toDataURL('image/jpeg', 0.95);
        pdf.addImage(imgData, 'JPEG', 0, 0, canvasWidth, pageHeight);
      }

      const pdfBlob = pdf.output('blob');
      resolve({
        blob: pdfBlob,
        url: URL.createObjectURL(pdfBlob),
        name: newName
      });
    } catch (err) {
      reject(err);
    }
  });
}

/**
 * Parses simple Markdown into HTML
 */
export function convertMarkdownToHTML(markdown: string): string {
  let html = markdown;
  
  // Headers
  html = html.replace(/^### (.*$)/gim, '<h3 class="text-lg font-bold mt-4 mb-2 text-slate-800">$1</h3>');
  html = html.replace(/^## (.*$)/gim, '<h2 class="text-xl font-bold mt-6 mb-3 text-slate-800 border-b pb-1">$1</h2>');
  html = html.replace(/^# (.*$)/gim, '<h1 class="text-2xl font-black mt-8 mb-4 text-slate-900">$1</h1>');
  
  // Bold
  html = html.replace(/\*\*(.*)\*\*/gim, '<strong>$1</strong>');
  html = html.replace(/__(.*)__/gim, '<strong>$1</strong>');

  // Italic
  html = html.replace(/\*(.*)\*/gim, '<em>$1</em>');
  html = html.replace(/_(.*)_/gim, '<em>$1</em>');

  // Code blocks
  html = html.replace(/```([\s\S]*?)```/gim, '<pre class="bg-slate-100 p-3 rounded font-mono text-sm my-3 border border-slate-200 overflow-auto">$1</pre>');
  html = html.replace(/`([^`]+)`/gim, '<code class="bg-slate-100 px-1 py-0.5 rounded font-mono text-xs text-red-600">$1</code>');

  // Blockquotes
  html = html.replace(/^\s*>\s+(.*$)/gim, '<blockquote class="border-r-4 border-indigo-500 pr-4 pl-2 text-slate-600 italic my-4">$1</blockquote>');

  // Unordered lists
  html = html.replace(/^\s*[\-\*]\s+(.*$)/gim, '<li class="list-disc list-inside mr-4 mb-1 text-slate-700">$1</li>');
  // Wrap li items in ul (simplified list grouping)
  html = html.replace(/(<li.*?>.*?<\/li>)/gim, '<ul class="my-3">$1</ul>');
  // Fix nested ul elements that are side-by-side
  html = html.replace(/<\/ul>\s*<ul class="my-3">/gim, '');

  // Ordered lists
  html = html.replace(/^\s*\d+\.\s+(.*$)/gim, '<li class="list-decimal list-inside mr-4 mb-1 text-slate-700">$1</li>');
  html = html.replace(/(<li class="list-decimal.*?>.*?<\/li>)/gim, '<ol class="my-3">$1</ol>');
  html = html.replace(/<\/ol>\s*<ol class="my-3">/gim, '');

  // Links
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/gim, '<a href="$2" target="_blank" class="text-indigo-600 hover:underline">$1</a>');

  // Paragraphs (simplified: split by empty lines, wrap non-HTML blocks)
  const paragraphs = html.split('\n\n');
  const finished = paragraphs.map(p => {
    p = p.trim();
    if (!p) return '';
    if (p.startsWith('<h') || p.startsWith('<u') || p.startsWith('<o') || p.startsWith('<b') || p.startsWith('<p') || p.startsWith('<pre')) {
      return p;
    }
    return `<p class="leading-relaxed mb-4 text-slate-700">${p}</p>`;
  });

  return finished.join('\n');
}

/**
 * Parses simple HTML back into Markdown
 */
export function convertHTMLToMarkdown(html: string): string {
  let md = html;

  // Strip scripts/styles
  md = md.replace(/<script[\s\S]*?<\/script>/gi, '');
  md = md.replace(/<style[\s\S]*?<\/style>/gi, '');

  // Headers
  md = md.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '# $1\n');
  md = md.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, '## $1\n');
  md = md.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, '### $1\n');

  // Bold / Italic
  md = md.replace(/<(strong|b)[^>]*>([\s\S]*?)<\/\1>/gi, '**$2**');
  md = md.replace(/<(em|i)[^>]*>([\s\S]*?)<\/\1>/gi, '*$2*');

  // Code
  md = md.replace(/<pre[^>]*><code[^>]*>([\s\S]*?)<\/code><\/pre>/gi, '```\n$1\n```\n');
  md = md.replace(/<pre[^>]*>([\s\S]*?)<\/pre>/gi, '```\n$1\n```\n');
  md = md.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, '`$1`');

  // Lists
  md = md.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, '- $1');
  md = md.replace(/<\/?(ul|ol)[^>]*>/gi, '\n');

  // Links
  md = md.replace(/<a[^>]*href=["']([^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');

  // Remove other HTML tags
  md = md.replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, '$1\n\n');
  md = md.replace(/<br\s*\/?>/gi, '\n');
  md = md.replace(/<[^>]+>/g, '');

  // Clean double line breaks
  md = md.replace(/\n{3,}/g, '\n\n');

  return md.trim();
}

/**
 * CSV Parser supporting quotes and linebreaks inside cells
 */
export function parseCSV(csvText: string): string[][] {
  const result: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          cell += '"';
          i++; // Skip double quote
        } else {
          inQuotes = false;
        }
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ',') {
        row.push(cell.trim());
        cell = '';
      } else if (char === '\r' || char === '\n') {
        row.push(cell.trim());
        if (row.length > 0 && (row.length > 1 || row[0] !== '')) {
          result.push(row);
        }
        row = [];
        cell = '';
        if (char === '\r' && nextChar === '\n') {
          i++; // Skip \n
        }
      } else {
        cell += char;
      }
    }
  }

  if (cell || row.length > 0) {
    row.push(cell.trim());
    result.push(row);
  }

  return result;
}

/**
 * JSON Array to CSV string
 */
export function jsonToCSV(jsonObj: any[]): string {
  if (!Array.isArray(jsonObj) || jsonObj.length === 0) {
    return '';
  }

  const keys = Array.from(
    new Set(jsonObj.reduce((acc, item) => acc.concat(Object.keys(item)), []))
  ) as string[];

  const csvRows = [];
  // Add Header
  csvRows.push(keys.map(key => `"${key.replace(/"/g, '""')}"`).join(','));

  // Add Data Rows
  for (const item of jsonObj) {
    const values = keys.map(key => {
      const val = item[key] !== undefined && item[key] !== null ? item[key] : '';
      const strVal = typeof val === 'object' ? JSON.stringify(val) : String(val);
      return `"${strVal.replace(/"/g, '""')}"`;
    });
    csvRows.push(values.join(','));
  }

  return csvRows.join('\n');
}

/**
 * XML parsing and conversion to JSON Object
 */
export function xmlToJSON(xmlText: string): any {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
  
  // Check for parsing errors
  const parseError = xmlDoc.getElementsByTagName('parsererror');
  if (parseError.length > 0) {
    throw new Error('الملف ليس بتنسيق XML صالح أو يحتوي على أخطاء برمجية.');
  }

  const nodeToJSON = (node: Node): any => {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.nodeValue?.trim() || '';
    }

    if (node.nodeType === Node.ELEMENT_NODE) {
      const element = node as Element;
      const obj: Record<string, any> = {};

      // Add Attributes with @ prefix
      if (element.attributes.length > 0) {
        for (let i = 0; i < element.attributes.length; i++) {
          const attr = element.attributes[i];
          obj[`@${attr.name}`] = attr.value;
        }
      }

      // Process children
      if (element.childNodes.length > 0) {
        let textVal = '';
        const childrenList: Record<string, any[]> = {};
        let hasElements = false;

        for (let i = 0; i < element.childNodes.length; i++) {
          const child = element.childNodes[i];
          if (child.nodeType === Node.ELEMENT_NODE) {
            hasElements = true;
            const childName = child.nodeName;
            const childVal = nodeToJSON(child);
            if (!childrenList[childName]) {
              childrenList[childName] = [];
            }
            childrenList[childName].push(childVal);
          } else if (child.nodeType === Node.TEXT_NODE) {
            textVal += child.nodeValue?.trim() || '';
          }
        }

        if (hasElements) {
          // Collapse single element arrays or keep arrays
          for (const key in childrenList) {
            if (childrenList[key].length === 1) {
              obj[key] = childrenList[key][0];
            } else {
              obj[key] = childrenList[key];
            }
          }
        } else if (textVal) {
          if (Object.keys(obj).length > 0) {
            obj['#text'] = textVal;
          } else {
            return textVal;
          }
        }
      }

      return Object.keys(obj).length === 0 ? '' : obj;
    }
    return null;
  };

  const rootElement = xmlDoc.documentElement;
  const result: Record<string, any> = {};
  result[rootElement.nodeName] = nodeToJSON(rootElement);
  return result;
}

/**
 * JSON object to XML string
 */
export function jsonToXML(obj: any, rootName = 'root'): string {
  const toXML = (value: any, key: string): string => {
    if (value === null || value === undefined) {
      return `<${key}></${key}>`;
    }

    if (Array.isArray(value)) {
      return value.map(item => toXML(item, key)).join('\n');
    }

    if (typeof value === 'object') {
      let xml = `<${key}`;
      let attrs = '';
      let children = '';

      for (const prop in value) {
        if (prop.startsWith('@')) {
          attrs += ` ${prop.substring(1)}="${String(value[prop]).replace(/"/g, '&quot;')}"`;
        } else if (prop === '#text') {
          children += String(value[prop])
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
        } else {
          children += toXML(value[prop], prop);
        }
      }

      xml += attrs + '>';
      xml += children;
      xml += `</${key}>`;
      return xml;
    }

    const cleanVal = String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    return `<${key}>${cleanVal}</${key}>`;
  };

  return `<?xml version="1.0" encoding="UTF-8"?>\n${toXML(obj, rootName)}`;
}

/**
 * Main Conversion Function Orchestrator
 */
/**
 * Main Conversion Function Orchestrator
 */
export async function performConversion(
  fileItem: { file: File; targetExtension: string }
): Promise<{ blob: Blob; url: string; name: string }> {
  const { file, targetExtension } = fileItem;
  let sourceExt = file.name.split('.').pop()?.toLowerCase() || '';

  // 0. Treat code files exactly like txt files for seamless text extraction and printing
  const codeExtensions = ['js', 'ts', 'py', 'c', 'cpp', 'cs', 'go', 'rs', 'php', 'css', 'sql', 'sh'];
  const isCodeFile = codeExtensions.includes(sourceExt);
  if (isCodeFile) {
    sourceExt = 'txt';
  }

  // 0.1 PDF Conversions (fully offline client-side support)
  if (sourceExt === 'pdf') {
    if (targetExtension === 'txt') {
      const text = await convertPdfToText(file);
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
    if (targetExtension === 'html') {
      const text = await convertPdfToText(file);
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const paragraphs = text.split('\n').map(p => p.trim() ? `<p style="margin-bottom: 1rem; line-height: 1.8;">${p}</p>` : '<br/>').join('\n');
      const htmlContent = `
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>${file.name}</title>
  <style>
    body { font-family: system-ui, sans-serif; padding: 2rem; max-width: 800px; margin: 0 auto; color: #1e293b; background-color: #f8fafc; }
    .card { background: white; padding: 2.5rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05); border: 1px solid #e2e8f0; }
  </style>
</head>
<body>
  <div class="card">
    <h2>محتوى المستند: ${originalBaseName}</h2>
    <hr style="margin: 1.5rem 0; border: 0; border-top: 1px solid #e2e8f0;"/>
    <div style="white-space: pre-wrap;">${paragraphs}</div>
  </div>
</body>
</html>`;
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.html` };
    }
    if (targetExtension === 'docx') {
      const text = await convertPdfToText(file);
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const docxBlob = await generateDocx(text);
      return { blob: docxBlob, url: URL.createObjectURL(docxBlob), name: `${originalBaseName}.docx` };
    }
    if (targetExtension === 'xlsx') {
      const text = await convertPdfToText(file);
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const rows = parseTextToRows(text);
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'PDF Text Data');
      const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.xlsx` };
    }
    if (targetExtension === 'png' || targetExtension === 'jpg' || targetExtension === 'jpeg') {
      return await convertPdfToImage(file, targetExtension);
    }
  }

  // 0.2 Word DOCX Conversions (fully offline client-side support)
  if (sourceExt === 'docx') {
    const text = await convertDocxToText(file);
    const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
    
    if (targetExtension === 'txt') {
      const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
    if (targetExtension === 'pdf') {
      return await convertTextToPDF(text, file.name);
    }
    if (targetExtension === 'md') {
      const mdContent = `# ${originalBaseName}\n\n${text}`;
      const blob = new Blob([mdContent], { type: 'text/markdown;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.md` };
    }
    if (targetExtension === 'html') {
      const paragraphs = text.split('\n').map(p => p.trim() ? `<p style="margin-bottom: 1rem; line-height: 1.8;">${p}</p>` : '<br/>').join('\n');
      const htmlContent = `
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>${file.name}</title>
  <style>
    body { font-family: system-ui, sans-serif; padding: 2rem; max-width: 800px; margin: 0 auto; color: #1e293b; background-color: #f8fafc; }
    .card { background: white; padding: 2.5rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05); border: 1px solid #e2e8f0; }
  </style>
</head>
<body>
  <div class="card">
    <h2>مستند وورد: ${originalBaseName}</h2>
    <hr style="margin: 1.5rem 0; border: 0; border-top: 1px solid #e2e8f0;"/>
    <div style="white-space: pre-wrap;">${paragraphs}</div>
  </div>
</body>
</html>`;
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.html` };
    }
  }

  // 0.3 Excel Sheets Conversions (fully offline client-side support)
  if (sourceExt === 'xlsx' || sourceExt === 'xls') {
    return await convertExcel(file, targetExtension);
  }

  // 1. Image conversions
  const imageExtensions = ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'svg', 'gif'];
  if (imageExtensions.includes(sourceExt)) {
    return await convertImage(file, targetExtension);
  }

  // 2. Text conversions
  if (sourceExt === 'txt') {
    const textContent = await readFileAsText(file);
    if (targetExtension === 'pdf') {
      return await convertTextToPDF(textContent, file.name);
    }
    if (targetExtension === 'html') {
      const htmlContent = `
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>${file.name}</title>
  <style>
    body { font-family: system-ui, sans-serif; padding: 2rem; max-width: 800px; margin: 0 auto; line-height: 1.6; color: #1e293b; }
    pre { background: #f1f5f9; padding: 1rem; border-radius: 4px; overflow-x: auto; }
  </style>
</head>
<body>
  <pre>${textContent.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>
</body>
</html>`;
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.html` };
    }
    if (targetExtension === 'md') {
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const mdContent = `# ${originalBaseName}\n\n${textContent}`;
      const blob = new Blob([mdContent], { type: 'text/markdown;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.md` };
    }
  }

  // 3. Markdown conversions
  if (sourceExt === 'md') {
    const textContent = await readFileAsText(file);
    if (targetExtension === 'html') {
      const parsedHTML = convertMarkdownToHTML(textContent);
      const htmlContent = `
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>${file.name}</title>
  <style>
    body { font-family: system-ui, sans-serif; padding: 2rem; max-width: 800px; margin: 0 auto; line-height: 1.7; color: #1e293b; background-color: #fafafa; }
    .content-card { background: white; padding: 2.5rem; border-radius: 8px; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); border: 1px solid #e2e8f0; }
  </style>
</head>
<body>
  <div class="content-card">
    ${parsedHTML}
  </div>
</body>
</html>`;
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.html` };
    }
    if (targetExtension === 'pdf') {
      const htmlText = convertMarkdownToHTML(textContent);
      // Strip HTML tags for simple PDF print representation
      const plainText = htmlText.replace(/<[^>]+>/g, '').replace(/\n\s*\n/g, '\n\n');
      return await convertTextToPDF(plainText, file.name);
    }
    if (targetExtension === 'txt') {
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
  }

  // 4. HTML Conversions
  if (sourceExt === 'html') {
    const textContent = await readFileAsText(file);
    if (targetExtension === 'md') {
      const mdText = convertHTMLToMarkdown(textContent);
      const blob = new Blob([mdText], { type: 'text/markdown;charset=utf-8' });
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.md` };
    }
    if (targetExtension === 'txt') {
      const plainText = textContent.replace(/<[^>]+>/g, '');
      const blob = new Blob([plainText], { type: 'text/plain;charset=utf-8' });
      const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
    if (targetExtension === 'pdf') {
      const plainText = textContent.replace(/<[^>]+>/g, '');
      return await convertTextToPDF(plainText, file.name);
    }
  }

  // 5. CSV Conversions
  if (sourceExt === 'csv') {
    const textContent = await readFileAsText(file);
    const parsedRows = parseCSV(textContent);
    
    // Create simple JSON structure
    const headers = parsedRows[0] || [];
    const jsonData = [];
    for (let r = 1; r < parsedRows.length; r++) {
      const row = parsedRows[r];
      const obj: Record<string, string> = {};
      headers.forEach((h, idx) => {
        obj[h || `column_${idx + 1}`] = row[idx] || '';
      });
      jsonData.push(obj);
    }

    const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;

    if (targetExtension === 'json') {
      const jsonStr = JSON.stringify(jsonData, null, 2);
      const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.json` };
    }
    if (targetExtension === 'xml') {
      const xmlStr = jsonToXML(jsonData, 'rows');
      const blob = new Blob([xmlStr], { type: 'application/xml;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.xml` };
    }
    if (targetExtension === 'yaml') {
      const yamlStr = jsYaml.dump(jsonData);
      const blob = new Blob([yamlStr], { type: 'text/yaml;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.yaml` };
    }
    if (targetExtension === 'txt') {
      const formattedText = parsedRows.map(row => row.join('\t')).join('\n');
      const blob = new Blob([formattedText], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
    if (targetExtension === 'xlsx') {
      // Create a XLSX file from CSV rows using SheetJS
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(parsedRows);
      XLSX.utils.book_append_sheet(wb, ws, 'CSV Data');
      const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.xlsx` };
    }
  }

  // 6. JSON Conversions
  if (sourceExt === 'json') {
    const textContent = await readFileAsText(file);
    const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
    
    let parsedObj: any;
    try {
      parsedObj = JSON.parse(textContent);
    } catch (e) {
      throw new Error('ملف JSON غير صالح. يرجى التحقق من صياغته.');
    }

    if (targetExtension === 'csv') {
      const dataArr = Array.isArray(parsedObj) ? parsedObj : [parsedObj];
      const csvStr = jsonToCSV(dataArr);
      const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.csv` };
    }
    if (targetExtension === 'yaml') {
      const yamlStr = jsYaml.dump(parsedObj);
      const blob = new Blob([yamlStr], { type: 'text/yaml;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.yaml` };
    }
    if (targetExtension === 'xml') {
      const xmlStr = jsonToXML(parsedObj, 'root');
      const blob = new Blob([xmlStr], { type: 'application/xml;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.xml` };
    }
    if (targetExtension === 'txt') {
      const formattedText = JSON.stringify(parsedObj, null, 4);
      const blob = new Blob([formattedText], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
    if (targetExtension === 'xlsx') {
      const dataArr = Array.isArray(parsedObj) ? parsedObj : [parsedObj];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(dataArr);
      XLSX.utils.book_append_sheet(wb, ws, 'JSON Data');
      const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.xlsx` };
    }
  }

  // 7. YAML Conversions
  if (sourceExt === 'yaml' || sourceExt === 'yml') {
    const textContent = await readFileAsText(file);
    const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
    
    let parsedObj: any;
    try {
      parsedObj = jsYaml.load(textContent);
    } catch (e) {
      throw new Error('ملف YAML غير صالح. يرجى التحقق من صياغته.');
    }

    if (targetExtension === 'json') {
      const jsonStr = JSON.stringify(parsedObj, null, 2);
      const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.json` };
    }
    if (targetExtension === 'xml') {
      const xmlStr = jsonToXML(parsedObj, 'root');
      const blob = new Blob([xmlStr], { type: 'application/xml;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.xml` };
    }
    if (targetExtension === 'txt') {
      const formattedText = JSON.stringify(parsedObj, null, 4);
      const blob = new Blob([formattedText], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
  }

  // 8. XML Conversions
  if (sourceExt === 'xml') {
    const textContent = await readFileAsText(file);
    const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
    
    let parsedObj: any;
    try {
      parsedObj = xmlToJSON(textContent);
    } catch (e) {
      throw new Error(e instanceof Error ? e.message : 'فشل في تحليل ملف XML');
    }

    if (targetExtension === 'json') {
      const jsonStr = JSON.stringify(parsedObj, null, 2);
      const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.json` };
    }
    if (targetExtension === 'csv') {
      // Find arrays inside parsed XML to convert to CSV
      let dataToConvert = parsedObj;
      const firstKey = Object.keys(parsedObj)[0];
      if (firstKey && typeof parsedObj[firstKey] === 'object') {
        dataToConvert = parsedObj[firstKey];
      }
      const dataArr = Array.isArray(dataToConvert) ? dataToConvert : [dataToConvert];
      const csvStr = jsonToCSV(dataArr);
      const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.csv` };
    }
    if (targetExtension === 'txt') {
      const formattedText = JSON.stringify(parsedObj, null, 4);
      const blob = new Blob([formattedText], { type: 'text/plain;charset=utf-8' });
      return { blob, url: URL.createObjectURL(blob), name: `${originalBaseName}.txt` };
    }
  }

  // 9. Fallback conversion for all other formats of the world (allows instant container rename & stream preservation)
  const originalBaseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
  const newName = `${originalBaseName}.${targetExtension}`;
  const fileBuffer = await readFileAsArrayBuffer(file);
  const fallbackBlob = new Blob([fileBuffer], { type: 'application/octet-stream' });
  return {
    blob: fallbackBlob,
    url: URL.createObjectURL(fallbackBlob),
    name: newName
  };
}

/**
 * Compresses multiple Files into a ZIP Archive
 */
export async function compressToZIP(
  files: File[],
  zipName = 'converted_files.zip'
): Promise<{ blob: Blob; url: string; name: string }> {
  const zip = new JSZip();
  
  for (const file of files) {
    const buffer = await readFileAsArrayBuffer(file);
    zip.file(file.name, buffer);
  }

  const blob = await zip.generateAsync({ type: 'blob' });
  return {
    blob,
    url: URL.createObjectURL(blob),
    name: zipName
  };
}

/**
 * Extracts a ZIP file and lists its contents
 */
export async function extractZIP(
  file: File
): Promise<Array<{ name: string; blob: Blob; url: string; size: number }>> {
  const zip = new JSZip();
  const loadedZip = await zip.loadAsync(file);
  const extractedFiles: Array<{ name: string; blob: Blob; url: string; size: number }> = [];

  const promises = Object.keys(loadedZip.files).map(async (fileName) => {
    const zipEntry = loadedZip.files[fileName];
    if (zipEntry.dir) return; // Skip directories

    const blob = await zipEntry.async('blob');
    extractedFiles.push({
      name: fileName,
      blob,
      url: URL.createObjectURL(blob),
      size: blob.size
    });
  });

  await Promise.all(promises);
  return extractedFiles;
}
