import { useState, useRef, DragEvent, ChangeEvent } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  UploadCloud,
  File,
  FileImage,
  FileText,
  FileCode,
  FolderArchive,
  ArrowLeftRight,
  Download,
  Trash2,
  X,
  Loader2,
  CheckCircle2,
  AlertTriangle,
  Lock,
  ChevronDown,
  FileSpreadsheet
} from 'lucide-react';
import { FileItem, ALL_WORLD_FORMATS, EXTENSION_MAPPED_TARGETS } from './types';
import { performConversion, extractZIP } from './utils/converters';

export default function App() {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [overallStatus, setOverallStatus] = useState<'idle' | 'converting' | 'completed' | 'failed'>('idle');
  const [isDragging, setIsDragging] = useState(false);

  // ZIP Extractor State (Simple & direct)
  const [zipExtractorMode, setZipExtractorMode] = useState(false);
  const [extractedFiles, setExtractedFiles] = useState<{ name: string; blob: Blob; url: string; size: number }[]>([]);
  const [extractedZipName, setExtractedZipName] = useState('');
  const [extracting, setExtracting] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const zipFileInputRef = useRef<HTMLInputElement>(null);

  // Format file sizes into human-readable units
  const formatSize = (bytes: number): string => {
    if (bytes === 0) return '0 بكسل';
    const k = 1024;
    const sizes = ['بايت', 'كيلوبايت', 'ميغابايت', 'جيجابايت'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  // Identify file category based on extension
  const getFileCategory = (ext: string): FileItem['category'] => {
    const images = ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'svg', 'gif', 'ico', 'tiff', 'psd', 'heic', 'ai', 'eps'];
    const docs = ['txt', 'pdf', 'html', 'md', 'docx', 'doc', 'rtf', 'epub', 'mobi', 'odt', 'pages'];
    const data = ['csv', 'json', 'yaml', 'yml', 'xml', 'xlsx', 'xls', 'ods'];
    const presentations = ['pptx', 'ppt', 'key', 'odp'];
    const audio = ['wav', 'mp3', 'aac', 'flac', 'ogg', 'm4a', 'wma', 'amr'];
    const video = ['mp4', 'mkv', 'avi', 'mov', 'webm', 'flv', 'wmv', '3gp'];
    const archive = ['zip', 'rar', 'tar', 'gz', '7z', 'iso'];
    const code = ['js', 'ts', 'py', 'c', 'cpp', 'cs', 'go', 'rs', 'php', 'css', 'sql', 'sh'];

    if (images.includes(ext)) return 'image';
    if (docs.includes(ext)) return 'document';
    if (data.includes(ext)) return 'data';
    if (presentations.includes(ext)) return 'presentation';
    if (audio.includes(ext)) return 'audio';
    if (video.includes(ext)) return 'video';
    if (archive.includes(ext)) return 'archive';
    if (code.includes(ext)) return 'code';
    return 'unknown';
  };

  // Add files to the queue
  const handleAddFiles = (fileList: FileList) => {
    const newItems: FileItem[] = [];

    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const extension = file.name.split('.').pop()?.toLowerCase() || '';
      const category = getFileCategory(extension);
      
      // Default target format selection based on EXTENSION_MAPPED_TARGETS
      const targets = EXTENSION_MAPPED_TARGETS[extension] || [];
      const defaultTarget = targets.length > 0 ? targets[0] : '';

      // Avoid duplicates
      if (files.some(f => f.file.name === file.name && f.file.size === file.size)) {
        continue;
      }

      newItems.push({
        id: crypto.randomUUID(),
        file,
        name: file.name,
        size: file.size,
        type: file.type,
        extension,
        category,
        status: 'idle',
        progress: 0,
        targetExtension: defaultTarget
      });
    }

    if (newItems.length > 0) {
      setFiles((prev) => [...prev, ...newItems]);
      setOverallStatus('idle');
    }
  };

  // Zip Extract Helper
  const handleZipExtract = async (file: File) => {
    setZipExtractorMode(true);
    setExtracting(true);
    try {
      const extracted = await extractZIP(file);
      setExtractedFiles(extracted);
    } catch (err) {
      alert('فشل في استخراج ملف الـ ZIP. قد يكون تالفاً أو غير مدعوم.');
    } finally {
      setExtracting(false);
    }
  };

  // Drag-and-drop Handlers
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files) {
      // Check if it's a zip file in zip extractor mode
      if (zipExtractorMode && e.dataTransfer.files[0]) {
        const file = e.dataTransfer.files[0];
        if (file.name.endsWith('.zip')) {
          setExtractedZipName(file.name);
          handleZipExtract(file);
          return;
        }
      }
      handleAddFiles(e.dataTransfer.files);
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      handleAddFiles(e.target.files);
    }
  };

  // Remove a single file from the conversion queue
  const removeFile = (id: string) => {
    setFiles((prev) => prev.filter((item) => item.id !== id));
  };

  // Update target extension for a specific file
  const updateTargetExtension = (id: string, targetExt: string) => {
    setFiles((prev) =>
      prev.map((item) =>
        item.id === id ? { ...item, targetExtension: targetExt } : item
      )
    );
  };

  // Execute conversion for a single file and trigger direct download immediately!
  const convertSingleFile = async (id: string) => {
    const fileItem = files.find((f) => f.id === id);
    if (!fileItem) return;

    // Set file state to converting
    setFiles((prev) =>
      prev.map((item) =>
        item.id === id
          ? { ...item, status: 'converting', progress: 30 }
          : item
      )
    );

    try {
      const progressInterval = setInterval(() => {
        setFiles((prev) =>
          prev.map((item) => {
            if (item.id === id && item.status === 'converting') {
              const nextProgress = Math.min(item.progress + 15, 90);
              return { ...item, progress: nextProgress };
            }
            return item;
          })
        );
      }, 100);

      // Perform actual conversion
      const result = await performConversion({
        file: fileItem.file,
        targetExtension: fileItem.targetExtension
      });

      clearInterval(progressInterval);

      setFiles((prev) =>
        prev.map((item) =>
          item.id === id
            ? {
                ...item,
                status: 'completed',
                progress: 100,
                convertedUrl: result.url,
                convertedName: result.name,
                convertedBlob: result.blob
              }
            : item
        )
      );

      // Direct Immediate Download!
      const link = document.createElement('a');
      link.href = result.url;
      link.download = result.name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

    } catch (err) {
      setFiles((prev) =>
        prev.map((item) =>
          item.id === id
            ? {
                ...item,
                status: 'failed',
                progress: 0,
                errorMsg: err instanceof Error ? err.message : 'خطأ في التحويل'
              }
            : item
        )
      );
    }
  };

  // Trigger conversion & individual download for all files in the queue
  const convertAllFiles = async () => {
    if (files.length === 0) return;
    setOverallStatus('converting');

    const pendingFiles = files.filter(f => 
      (f.status === 'idle' || f.status === 'failed') && 
      (EXTENSION_MAPPED_TARGETS[f.extension] || []).length > 0
    );
    if (pendingFiles.length === 0) {
      setOverallStatus('completed');
      return;
    }

    // Convert concurrently and trigger downloads
    await Promise.all(pendingFiles.map(f => convertSingleFile(f.id)));
    setOverallStatus('completed');
  };

  // Trigger input file dialog click
  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  const triggerZipFileInput = () => {
    zipFileInputRef.current?.click();
  };

  // Render appropriate Lucide Icon for a file category
  const getFileIcon = (category: FileItem['category'], className = 'w-6 h-6') => {
    switch (category) {
      case 'image':
        return <FileImage className={`${className} text-emerald-500`} />;
      case 'document':
        return <FileText className={`${className} text-blue-500`} />;
      case 'data':
        return <FileSpreadsheet className={`${className} text-teal-500`} />;
      case 'archive':
        return <FolderArchive className={`${className} text-amber-500`} />;
      case 'audio':
        return <File className={`${className} text-pink-500`} />;
      case 'video':
        return <File className={`${className} text-rose-500`} />;
      case 'presentation':
        return <FileText className={`${className} text-violet-500`} />;
      case 'code':
        return <FileCode className={`${className} text-indigo-500`} />;
      default:
        return <File className={`${className} text-slate-400`} />;
    }
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-900 flex flex-col font-sans selection:bg-indigo-500 selection:text-white" dir="rtl">
      
      {/* Top Beautiful Navbar */}
      <nav className="sticky top-0 z-50 bg-white/85 backdrop-blur-md border-b border-slate-100 py-4 px-6 md:px-12 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-600 p-2.5 rounded-2xl text-white shadow-lg shadow-indigo-600/20 flex items-center justify-center">
            <ArrowLeftRight className="w-5 h-5 animate-pulse" />
          </div>
          <div>
            <span className="font-extrabold text-xl tracking-wider bg-gradient-to-r from-indigo-600 to-violet-600 bg-clip-text text-transparent uppercase">
              Convix
            </span>
            <span className="text-[10px] font-mono block text-slate-400">محول الصيغ العصري السريع</span>
          </div>
        </div>

        {/* Simplest Switch Toggle */}
        <div className="flex items-center gap-2">
          <button
            id="switch-converter-mode-btn"
            onClick={() => {
              setZipExtractorMode(false);
              setExtractedFiles([]);
            }}
            className={`px-4 py-2 rounded-xl text-xs font-bold transition-all duration-200 cursor-pointer ${
              !zipExtractorMode
                ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/10'
                : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            محول الصيغ
          </button>
          <button
            id="switch-zip-mode-btn"
            onClick={() => {
              setZipExtractorMode(true);
              setFiles([]);
            }}
            className={`px-4 py-2 rounded-xl text-xs font-bold transition-all duration-200 cursor-pointer ${
              zipExtractorMode
                ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/10'
                : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            مستخرج ZIP
          </button>
        </div>
      </nav>

      <main className="flex-1 max-w-3xl w-full mx-auto px-4 py-8 flex flex-col gap-6 justify-center">
        
        {/* Simple Definition Section */}
        <header className="text-center max-w-xl mx-auto mb-2">
          <h1 className="text-3xl font-black text-slate-950 tracking-tight leading-tight">
            {zipExtractorMode ? 'مستخرج ملفات ZIP البسيط' : 'محول صيغ الملفات الفوري'}
          </h1>
          <p className="mt-2 text-sm text-slate-500 leading-relaxed font-normal">
            {zipExtractorMode
              ? 'افتح واستخرج محتويات ملفات الأرشيف والـ ZIP محلياً بضغطة زر واحدة بكل سهولة وأمان.'
              : 'اختر أي ملف من جهازك، حدد الصيغة التي تريد التحويل إليها، وسيتم تحميل الملف بالصيغة الجديدة فوراً دون رفع أي بيانات.'}
          </p>
        </header>

        {/* MAIN WORKSPACE */}
        {!zipExtractorMode ? (
          <div className="flex flex-col gap-5">
            
            {/* Elegant Selector Box / Drag and Drop */}
            <div
              id="dropzone"
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={triggerFileInput}
              className={`relative border-2 border-dashed rounded-3xl p-8 text-center cursor-pointer transition-all duration-200 flex flex-col items-center justify-center group bg-white ${
                isDragging
                  ? 'border-indigo-600 bg-indigo-50/50 scale-[0.99] shadow-inner'
                  : 'border-slate-200 hover:border-indigo-400 hover:shadow-lg'
              }`}
            >
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileChange}
                multiple
                className="hidden"
              />

              <div className="bg-indigo-50 text-indigo-600 p-4 rounded-full mb-3 group-hover:scale-105 transition-all duration-200">
                <UploadCloud className="w-8 h-8" />
              </div>

              <h3 className="text-sm font-extrabold text-slate-800">
                اضغط هنا لتحديد الملف أو اسحبه وأفلته في هذا المربع
              </h3>
              <p className="text-xs text-slate-400 mt-1">
                يدعم جميع صيغ الصور، المستندات، الجداول، الفيديو، الصوتيات والملفات البرمجية
              </p>
            </div>

            {/* Uploaded Files Queue / Selection List */}
            {files.length > 0 && (
              <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-5 flex flex-col gap-3">
                <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-1">
                  <h2 className="font-extrabold text-slate-800 text-sm flex items-center gap-2">
                    <span>الملفات المختارة</span>
                    <span className="bg-slate-100 text-slate-600 text-xs px-2.5 py-0.5 rounded-full font-bold">
                      {files.length}
                    </span>
                  </h2>
                  <button
                    id="clear-all-btn"
                    onClick={() => setFiles([])}
                    className="text-xs text-rose-500 font-bold hover:underline flex items-center gap-1 cursor-pointer"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    مسح القائمة
                  </button>
                </div>

                <div className="flex flex-col gap-2.5 max-h-[350px] overflow-y-auto">
                  <AnimatePresence initial={false}>
                    {files.map((item) => (
                      <motion.div
                        layout
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, scale: 0.95 }}
                        key={item.id}
                        className="flex flex-col sm:flex-row items-start sm:items-center justify-between p-3.5 rounded-2xl bg-slate-50 border border-slate-100 hover:border-indigo-100 transition-all duration-150 gap-3"
                      >
                        <div className="flex items-center gap-3 overflow-hidden w-full sm:w-auto">
                          <div className="p-2 bg-white rounded-xl border border-slate-100 flex-shrink-0">
                            {getFileIcon(item.category, 'w-5 h-5')}
                          </div>
                          <div className="overflow-hidden">
                            <h4 className="font-bold text-slate-800 text-xs truncate max-w-[180px] sm:max-w-[220px]" title={item.name}>
                              {item.name}
                            </h4>
                            <div className="flex items-center gap-2 mt-0.5 text-[10px] text-slate-400 font-semibold">
                              <span>{formatSize(item.size)}</span>
                              <span>•</span>
                              <span className="uppercase text-[9px] bg-slate-200/60 px-1 py-0.2 rounded font-mono text-slate-600">
                                {item.extension}
                              </span>
                            </div>
                            {item.status === 'failed' && item.errorMsg && (
                              <p className="text-[10px] text-rose-500 font-semibold mt-1">
                                السبب: {item.errorMsg}
                              </p>
                            )}
                          </div>
                        </div>

                        {/* Format selector and action button */}
                        <div className="flex items-center gap-2 w-full sm:w-auto justify-between sm:justify-end">
                          {item.status === 'idle' && (
                            <>
                              {(EXTENSION_MAPPED_TARGETS[item.extension] || []).length > 0 ? (
                                <div className="flex items-center gap-1 bg-white rounded-xl border border-slate-200 px-2.5 py-1 shadow-sm">
                                  <span className="text-[10px] font-bold text-slate-400">إلى</span>
                                  <div className="relative flex items-center">
                                    <select
                                      id={`format-select-${item.id}`}
                                      value={item.targetExtension}
                                      onChange={(e) => updateTargetExtension(item.id, e.target.value)}
                                      className="appearance-none bg-transparent pr-3 pl-5 text-xs font-bold text-indigo-600 outline-none cursor-pointer text-left font-mono"
                                    >
                                      {(EXTENSION_MAPPED_TARGETS[item.extension] || []).map((tgt) => (
                                        <option key={tgt} value={tgt} className="text-slate-900 font-semibold font-mono">
                                          {tgt.toUpperCase()}
                                        </option>
                                      ))}
                                    </select>
                                    <ChevronDown className="w-3.5 h-3.5 text-slate-400 absolute left-0 pointer-events-none" />
                                  </div>
                                </div>
                              ) : (
                                <span className="text-[11px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-xl">
                                  للاستخراج فقط (ZIP)
                                </span>
                              )}
                            </>
                          )}

                          {/* Trigger single file conversion */}
                          <div className="flex items-center gap-1.5">
                            {item.status === 'converting' && (
                              <div className="flex items-center gap-1.5 text-indigo-600 bg-indigo-50/80 px-2.5 py-1.5 rounded-xl text-xs font-bold animate-pulse">
                                <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                <span>جاري التحويل...</span>
                              </div>
                            )}

                            {item.status === 'completed' && (
                              <div className="flex items-center gap-1.5 text-emerald-600 bg-emerald-50 px-2.5 py-1.5 rounded-xl text-xs font-bold">
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                <span>تم الحفظ والتنزيل</span>
                              </div>
                            )}

                            {item.status === 'failed' && (
                              <div className="flex items-center gap-1.5 text-rose-600 bg-rose-50 px-2.5 py-1.5 rounded-xl text-xs font-bold">
                                <AlertTriangle className="w-3.5 h-3.5" />
                                <span>فشل</span>
                              </div>
                            )}

                            {item.status === 'idle' && (EXTENSION_MAPPED_TARGETS[item.extension] || []).length > 0 && (
                              <button
                                id={`convert-btn-${item.id}`}
                                onClick={() => convertSingleFile(item.id)}
                                className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-3.5 py-1.5 rounded-xl shadow-sm transition-all duration-150 cursor-pointer flex items-center gap-1"
                              >
                                <span>تحويل وتنزيل</span>
                                <Download className="w-3 h-3" />
                              </button>
                            )}

                            {item.status === 'completed' && item.convertedUrl && (
                              <a
                                id={`redownload-btn-${item.id}`}
                                href={item.convertedUrl}
                                download={item.convertedName}
                                className="bg-emerald-600 hover:bg-emerald-700 text-white p-2 rounded-xl transition-all duration-150 flex items-center justify-center"
                                title="تحميل مجدداً"
                              >
                                <Download className="w-3.5 h-3.5" />
                              </a>
                            )}

                            <button
                              id={`remove-btn-${item.id}`}
                              onClick={() => removeFile(item.id)}
                              className="text-slate-400 hover:text-rose-500 hover:bg-rose-50 p-2 rounded-xl transition-all duration-150 cursor-pointer"
                              title="إزالة من القائمة"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      </motion.div>
                    ))}
                  </AnimatePresence>
                </div>

                {/* Batch Button */}
                {files.some(f => (f.status === 'idle' || f.status === 'failed') && (EXTENSION_MAPPED_TARGETS[f.extension] || []).length > 0) && (
                  <button
                    id="convert-all-btn"
                    onClick={convertAllFiles}
                    disabled={overallStatus === 'converting'}
                    className="w-full mt-2 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-xs flex items-center justify-center gap-2 transition-all duration-150 cursor-pointer shadow-sm shadow-indigo-600/10"
                  >
                    {overallStatus === 'converting' ? (
                      <>
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        <span>جاري تحويل وتنزيل جميع الملفات...</span>
                      </>
                    ) : (
                      <>
                        <ArrowLeftRight className="w-3.5 h-3.5" />
                        <span>تحويل وتنزيل كافة الملفات ({files.filter(f => f.status === 'idle' && (EXTENSION_MAPPED_TARGETS[f.extension] || []).length > 0).length})</span>
                      </>
                    )}
                  </button>
                )}
              </div>
            )}
          </div>
        ) : (
          
          /* ZIP EXTRACTOR WORKSPACE */
          <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6 flex flex-col gap-5">
            {extractedFiles.length === 0 ? (
              <div
                id="zip-dropzone"
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={triggerZipFileInput}
                className={`border-2 border-dashed rounded-3xl p-8 text-center cursor-pointer transition-all duration-200 flex flex-col items-center justify-center bg-slate-50 hover:bg-white ${
                  isDragging ? 'border-amber-500 bg-amber-50/50' : 'border-slate-200 hover:border-amber-400'
                }`}
              >
                <input
                  type="file"
                  ref={zipFileInputRef}
                  onChange={(e) => {
                    if (e.target.files && e.target.files[0]) {
                      const file = e.target.files[0];
                      setExtractedZipName(file.name);
                      handleZipExtract(file);
                    }
                  }}
                  accept=".zip"
                  className="hidden"
                />

                <FolderArchive className="w-10 h-10 text-amber-500 mb-3 animate-bounce" />
                <h3 className="text-sm font-extrabold text-slate-800">اضغط هنا لتحديد ملف الـ ZIP أو اسحبه هنا</h3>
                <p className="text-xs text-slate-400 mt-1">سنقوم بتحليله وعرض الملفات الموجودة بداخله فوراً للاستخراج</p>
              </div>
            ) : (
              /* Extracted Files List */
              <div className="flex flex-col gap-4">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center bg-slate-50 p-4 rounded-2xl border border-slate-100 gap-2">
                  <div className="text-xs font-semibold text-slate-600">
                    <span>ملف الأرشيف: </span>
                    <span className="font-bold text-slate-800">{extractedZipName}</span>
                  </div>
                  <div className="text-xs font-semibold text-slate-600">
                    <span>عدد الملفات: </span>
                    <span className="font-bold text-amber-600">{extractedFiles.length} ملف جاهز للتنزيل</span>
                  </div>
                </div>

                <div className="flex flex-col gap-2 max-h-[300px] overflow-y-auto pr-1">
                  {extractedFiles.map((file, idx) => {
                    const ext = file.name.split('.').pop()?.toLowerCase() || '';
                    const cat = getFileCategory(ext);

                    return (
                      <div key={idx} className="flex items-center justify-between p-3 bg-slate-50 border border-slate-100 rounded-2xl gap-2">
                        <div className="flex items-center gap-3 overflow-hidden">
                          <div className="p-1.5 bg-white rounded-lg border border-slate-100 flex-shrink-0">
                            {getFileIcon(cat, 'w-4 h-4')}
                          </div>
                          <div className="overflow-hidden">
                            <h4 className="font-bold text-slate-800 text-xs truncate max-w-[200px]" title={file.name}>
                              {file.name}
                            </h4>
                            <span className="text-[10px] text-slate-400 font-semibold">{formatSize(file.size)}</span>
                          </div>
                        </div>

                        <a
                          id={`download-extracted-btn-${idx}`}
                          href={file.url}
                          download={file.name}
                          className="bg-amber-500 hover:bg-amber-600 text-white px-3 py-1.5 rounded-lg font-bold text-xs flex items-center gap-1"
                        >
                          <Download className="w-3 h-3" />
                          <span>تحميل</span>
                        </a>
                      </div>
                    );
                  })}
                </div>

                <div className="flex justify-end gap-2.5 mt-2 pt-3 border-t border-slate-100">
                  <button
                    id="reset-extractor-btn"
                    onClick={() => {
                      setExtractedFiles([]);
                      setExtractedZipName('');
                    }}
                    className="border border-slate-200 hover:bg-slate-50 text-slate-600 font-bold text-xs px-3.5 py-2 rounded-xl cursor-pointer"
                  >
                    استخراج ملف آخر
                  </button>
                  <button
                    id="download-all-extracted-btn"
                    onClick={() => {
                      extractedFiles.forEach((file) => {
                        const link = document.createElement('a');
                        link.href = file.url;
                        link.download = file.name;
                        link.click();
                      });
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-3.5 py-2 rounded-xl shadow-md cursor-pointer"
                  >
                    تنزيل جميع الملفات
                  </button>
                </div>
              </div>
            )}

            {extracting && (
              <div className="flex flex-col items-center justify-center p-8 gap-3">
                <Loader2 className="w-7 h-7 text-amber-500 animate-spin" />
                <span className="text-xs font-bold text-slate-600">جاري قراءة ملف الأرشيف واستخراج المحتويات محلياً...</span>
              </div>
            )}
          </div>
        )}

        {/* Security & Encryption Info Box */}
        <div className="bg-slate-50 border border-slate-100 rounded-3xl p-5 flex gap-4 mt-2">
          <div className="bg-emerald-50 text-emerald-600 p-2.5 rounded-2xl flex-shrink-0 self-start">
            <Lock className="w-5 h-5" />
          </div>
          <div>
            <h4 className="font-extrabold text-xs text-slate-800">بياناتك مشفرة ومحمية بالكامل</h4>
            <p className="text-[11px] text-slate-500 leading-relaxed mt-1 font-normal">
              جميع عمليات التحويل واستخراج الملفات تتم بالكامل محلياً داخل جهازك ومتصفحك مباشرة عبر تقنيات WebAssembly والـ Canvas. لا نقوم برفع أو إرسال ملفاتك لأي خادم خارجي، مما يضمن سرية تامة وأماناً مطلقاً لخصوصيتك.
            </p>
          </div>
        </div>

      </main>

      {/* Footer copyright */}
      <footer className="bg-white border-t border-slate-100 py-5 text-center text-xs text-slate-400 font-semibold mt-auto">
        <p>© 2026 Convix. جميع الحقوق محفوظة. تم تصميمه بكل إتقان لخدمتك بأمان.</p>
      </footer>
    </div>
  );
}
