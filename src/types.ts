export interface FileItem {
  id: string;
  file: File;
  name: string;
  size: number;
  type: string;
  extension: string;
  category: 'image' | 'document' | 'data' | 'archive' | 'audio' | 'video' | 'presentation' | 'code' | 'unknown';
  status: 'idle' | 'converting' | 'completed' | 'failed';
  progress: number;
  targetExtension: string;
  errorMsg?: string;
  convertedUrl?: string;
  convertedName?: string;
  convertedBlob?: Blob;
}

export type FileCategory = 'all' | 'image' | 'document' | 'data' | 'archive' | 'audio' | 'video' | 'code';

export interface FormatGroup {
  label: string;
  formats: string[];
}

export const CATEGORY_LABELS: Record<FileCategory, string> = {
  all: 'الكل',
  image: 'صور',
  document: 'مستندات',
  data: 'بيانات وجداول',
  archive: 'أرشيف وضغط',
  audio: 'صوتيات',
  video: 'فيديو مرئي',
  code: 'أكواد برمجية'
};

// All world-famous extensions grouped by category
export const ALL_WORLD_FORMATS: FormatGroup[] = [
  {
    label: 'الصور والرسومات',
    formats: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'ico', 'svg', 'tiff', 'psd', 'heic', 'ai', 'eps']
  },
  {
    label: 'المستندات والكتب الرقمية',
    formats: ['pdf', 'docx', 'doc', 'txt', 'rtf', 'epub', 'mobi', 'odt', 'pages']
  },
  {
    label: 'البيانات والجداول',
    formats: ['xlsx', 'xls', 'csv', 'json', 'xml', 'yaml', 'yml', 'ods']
  },
  {
    label: 'العروض التقديمية',
    formats: ['pptx', 'ppt', 'key', 'odp']
  },
  {
    label: 'الملفات الصوتية',
    formats: ['mp3', 'wav', 'm4a', 'aac', 'flac', 'ogg', 'wma', 'amr']
  },
  {
    label: 'ملفات الفيديو والمرئيات',
    formats: ['mp4', 'mkv', 'avi', 'mov', 'webm', 'flv', 'wmv', '3gp']
  },
  {
    label: 'الملفات المضغوطة والأرشيف',
    formats: ['zip', 'rar', '7z', 'tar', 'gz', 'iso']
  },
  {
    label: 'لغات البرمجة والأكواد',
    formats: ['js', 'ts', 'py', 'c', 'cpp', 'cs', 'go', 'rs', 'php', 'html', 'css', 'sql', 'sh']
  }
];

export const EXTENSION_MAPPED_TARGETS: Record<string, string[]> = {
  // Images (fully supported via HTML Canvas & jsPDF)
  png: ['jpg', 'webp', 'bmp', 'pdf'],
  jpg: ['png', 'webp', 'bmp', 'pdf'],
  jpeg: ['png', 'webp', 'bmp', 'pdf'],
  webp: ['png', 'jpg', 'bmp', 'pdf'],
  bmp: ['png', 'jpg', 'webp', 'pdf'],
  svg: ['png', 'jpg', 'webp', 'pdf'],
  gif: ['png', 'jpg', 'webp', 'pdf'],

  // PDF Documents (fully supported via PDF.js rendering & text parsing)
  pdf: ['docx', 'xlsx', 'txt', 'html', 'png', 'jpg'],

  // Word Documents (fully supported via JSZip & DOMParser extraction)
  docx: ['txt', 'pdf', 'html', 'md'],

  // Excel Sheets (fully supported via SheetJS parser & converters)
  xlsx: ['csv', 'json', 'html', 'txt'],
  xls: ['csv', 'json', 'html', 'txt'],

  // Documents / Text
  txt: ['pdf', 'html', 'md'],
  md: ['html', 'pdf', 'txt'],
  html: ['pdf', 'txt', 'md'],

  // Data
  csv: ['xlsx', 'json', 'xml', 'yaml', 'txt'],
  json: ['csv', 'yaml', 'xml', 'txt', 'xlsx'],
  xml: ['json', 'csv', 'txt'],
  yaml: ['json', 'xml', 'txt'],
  yml: ['json', 'xml', 'txt'],

  // Code files (handled elegantly as plain text)
  js: ['txt', 'pdf', 'html'],
  ts: ['txt', 'pdf', 'html'],
  py: ['txt', 'pdf', 'html'],
  c: ['txt', 'pdf', 'html'],
  cpp: ['txt', 'pdf', 'html'],
  cs: ['txt', 'pdf', 'html'],
  go: ['txt', 'pdf', 'html'],
  rs: ['txt', 'pdf', 'html'],
  php: ['txt', 'pdf', 'html'],
  css: ['txt', 'pdf', 'html'],
  sql: ['txt', 'pdf', 'html'],
  sh: ['txt', 'pdf', 'html']
};

