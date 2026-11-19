import React, { useMemo, useRef, useState } from 'react';
import {
  Check,
  Code2,
  Copy,
  Download,
  FileSpreadsheet,
  FolderArchive,
  Search,
  Upload,
} from 'lucide-react';
import scheduleSplitterPy from '../python_app/schedule_splitter.py?raw';
import generateSamplePy from '../python_app/generate_sample_schedule.py?raw';
import requirementsTxt from '../python_app/requirements.txt?raw';
import {
  exportAllSubjectsZip,
  exportSubjectWorkbookBuffer,
  generateSampleScheduleBuffer,
  getSubjectSessions,
  parseScheduleBuffer,
  type ParseResult,
} from './utils/scheduleEngine';

type ActiveTab = 'processor' | 'preview' | 'python' | 'architecture';
type PythonFileKey = 'schedule_splitter.py' | 'generate_sample_schedule.py' | 'requirements.txt';

function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

const PYTHON_FILES: Record<PythonFileKey, { label: string; content: string; description: string }> = {
  'schedule_splitter.py': {
    label: 'schedule_splitter.py',
    description: 'Chương trình chính: Quét động khối tuần, Regex tiêu đề, GUI Tkinter và xuất Excel bằng openpyxl.',
    content: scheduleSplitterPy,
  },
  'generate_sample_schedule.py': {
    label: 'generate_sample_schedule.py',
    description: 'Script tạo file Schedule_17_Weeks.xlsx và Schedule_21_Weeks.xlsx mẫu để kiểm thử.',
    content: generateSamplePy,
  },
  'requirements.txt': {
    label: 'requirements.txt',
    description: 'Thư viện Python cần cài đặt trước khi chạy ứng dụng Tkinter.',
    content: requirementsTxt,
  },
};

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [activeTab, setActiveTab] = useState<ActiveTab>('processor');
  const [fileName, setFileName] = useState<string | null>(null);
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [lecturerMap, setLecturerMap] = useState<Record<string, string>>({});
  const [selectedSubject, setSelectedSubject] = useState<string>('');
  const [previewSearch, setPreviewSearch] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);

  const [activePyFile, setActivePyFile] = useState<PythonFileKey>('schedule_splitter.py');
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [showIgnoredNotes, setShowIgnoredNotes] = useState<boolean>(false);

  async function processExcelBuffer(buffer: ArrayBuffer, displayFileName: string) {
    setIsLoading(true);
    setErrorMessage(null);
    setStatusMessage(null);
    try {
      const result = await parseScheduleBuffer(buffer);
      setParseResult(result);
      setFileName(displayFileName);
      const initialLecturers: Record<string, string> = {};
      result.subjects.forEach((subj) => {
        initialLecturers[subj] = '';
      });
      setLecturerMap(initialLecturers);
      setSelectedSubject(result.subjects[0] ?? '');
      setStatusMessage(
        `Đã quét động ${result.metadata.totalBlocks} khối (${result.metadata.totalWeeks} tuần), trích xuất ${result.records.length} buổi học thuộc ${result.subjects.length} môn.`
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Không thể đọc file Excel.';
      setErrorMessage(msg);
      setParseResult(null);
    } finally {
      setIsLoading(false);
    }
  }

  async function handleFileSelect(file: File | undefined) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setErrorMessage('Vui lòng chọn file định dạng .xlsx hợp lệ.');
      return;
    }
    const buffer = await file.arrayBuffer();
    await processExcelBuffer(buffer, file.name);
  }

  async function handleLoadSample(variant: '17w' | '21w') {
    setIsLoading(true);
    setErrorMessage(null);
    try {
      const is17 = variant === '17w';
      const buffer = await generateSampleScheduleBuffer({
        totalWeeks: is17 ? 17 : 21,
        isAfternoon: is17,
        className: is17 ? 'PC2204' : 'SE2102',
        room: is17 ? '107' : '304',
        term: is17 ? 'FALL 2026' : 'SPRING 2027',
        startDate: is17 ? new Date(2026, 7, 31) : new Date(2027, 0, 18),
      });
      const sampleName = is17 ? 'Schedule_PC2204_17Tuan.xlsx' : 'Schedule_SE2102_21Tuan.xlsx';
      await processExcelBuffer(buffer, sampleName);

      // Gợi ý tên giảng viên mẫu khi người dùng nạp dữ liệu thử nghiệm
      if (is17) {
        setLecturerMap({
          CSI104: 'Nguyễn Văn An',
          ENG101: 'Trần Thị Bích',
          MAT001: 'Lê Minh Khôi',
          PRO192: 'Phạm Hoàng Nam',
          WEB201: 'Đặng Thu Hà',
        });
      } else {
        setLecturerMap({
          CEA201: 'Vũ Quang Huy',
          MAE101: 'Lê Minh Khôi',
          NWC203: 'Trần Quốc Bảo',
          PRF192: 'Phạm Hoàng Nam',
          SSL101: 'Nguyễn Thị Lan',
        });
      }
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : 'Lỗi tạo dữ liệu mẫu.');
      setIsLoading(false);
    }
  }

  async function handleDownloadSampleExcel(variant: '17w' | '21w') {
    const is17 = variant === '17w';
    const buffer = await generateSampleScheduleBuffer({
      totalWeeks: is17 ? 17 : 21,
      isAfternoon: is17,
      className: is17 ? 'PC2204' : 'SE2102',
      room: is17 ? '107' : '304',
      term: is17 ? 'FALL 2026' : 'SPRING 2027',
      startDate: is17 ? new Date(2026, 7, 31) : new Date(2027, 0, 18),
    });
    const name = is17 ? 'Schedule_PC2204_17Tuan.xlsx' : 'Schedule_SE2102_21Tuan.xlsx';
    triggerDownload(
      new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      name
    );
  }

  async function handleExportSingleSubject(subject: string) {
    if (!parseResult) return;
    setIsExporting(true);
    try {
      const { filename, buffer } = await exportSubjectWorkbookBuffer(
        parseResult,
        subject,
        lecturerMap[subject] || ''
      );
      triggerDownload(
        new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
        filename
      );
      setStatusMessage(`Đã xuất file ${filename}.`);
    } finally {
      setIsExporting(false);
    }
  }

  async function handleExportAllZip() {
    if (!parseResult) return;
    setIsExporting(true);
    try {
      const { zipFilename, blob } = await exportAllSubjectsZip(parseResult, lecturerMap);
      triggerDownload(blob, zipFilename);
      setStatusMessage(
        `Đã xuất thành công ${parseResult.subjects.length} file Excel theo từng môn trong gói ${zipFilename}.`
      );
    } finally {
      setIsExporting(false);
    }
  }

  function handleCopyCode(key: PythonFileKey) {
    navigator.clipboard.writeText(PYTHON_FILES[key].content);
    setCopiedKey(key);
    window.setTimeout(() => setCopiedKey(null), 1800);
  }

  function handleDownloadPythonFile(key: PythonFileKey) {
    const blob = new Blob([PYTHON_FILES[key].content], { type: 'text/plain;charset=utf-8' });
    triggerDownload(blob, key);
  }

  const subjectSummaries = useMemo(() => {
    if (!parseResult) return [];
    return parseResult.subjects.map((subj) => {
      const sessions = getSubjectSessions(parseResult, subj, lecturerMap[subj] || '');
      return {
        subject: subj,
        count: sessions.length,
        firstDate: sessions[0]?.dateStr ?? '—',
        lastDate: sessions[sessions.length - 1]?.dateStr ?? '—',
        lecturer: lecturerMap[subj] ?? '',
      };
    });
  }, [parseResult, lecturerMap]);

  const previewSessions = useMemo(() => {
    if (!parseResult || !selectedSubject) return [];
    const list = getSubjectSessions(
      parseResult,
      selectedSubject,
      lecturerMap[selectedSubject] || ''
    );
    const query = previewSearch.trim().toLowerCase();
    if (!query) return list;
    return list.filter(
      (s) =>
        s.dateStr.toLowerCase().includes(query) ||
        String(s.slot).includes(query) ||
        s.lecturer.toLowerCase().includes(query) ||
        String(s.sessionNo).includes(query)
    );
  }, [parseResult, selectedSubject, lecturerMap, previewSearch]);

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx"
        className="hidden"
        onChange={(e) => {
          handleFileSelect(e.target.files?.[0]);
          e.target.value = '';
        }}
      />

      {/* Top Bar Contract: 3 zones separated by gap-8 */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between gap-8">
          {/* Zone 1: Single text wordmark */}
          <a
            href="#top"
            onClick={(e) => {
              e.preventDefault();
              setActiveTab('processor');
            }}
            className="font-display text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
          >
            Schedule Splitter
          </a>

          {/* Zone 2: Clean single-line navigation links */}
          <nav className="hidden md:flex items-center gap-7 text-sm font-medium">
            <button
              type="button"
              onClick={() => setActiveTab('processor')}
              className={`py-5 border-b-2 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'processor'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Xử lý Excel
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('preview')}
              className={`py-5 border-b-2 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'preview'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Xem trước lịch
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('python')}
              className={`py-5 border-b-2 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'python'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Mã nguồn Python
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('architecture')}
              className={`py-5 border-b-2 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'architecture'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Cơ chế quét động
            </button>
          </nav>

          {/* Zone 3: Single primary CTA */}
          <div className="flex items-center shrink-0">
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-700"
            >
              <Upload className="w-4 h-4" />
              Chọn file Schedule.xlsx
            </button>
          </div>
        </div>

        {/* Mobile navigation strip */}
        <div className="md:hidden border-t border-slate-100 px-4 flex items-center gap-4 overflow-x-auto text-xs font-medium">
          {(
            [
              ['processor', 'Xử lý Excel'],
              ['preview', 'Xem trước'],
              ['python', 'Python Tkinter'],
              ['architecture', 'Quét động'],
            ] as const
          ).map(([id, label]) => (
            <button
              key={id}
              type="button"
              onClick={() => setActiveTab(id)}
              className={`py-2.5 border-b-2 whitespace-nowrap shrink-0 ${
                activeTab === id
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600'
              }`}
            >
              {label}
            </button>
          ))}
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto px-6 py-8 space-y-6">
        {/* Status / Error Feedback */}
        {errorMessage && (
          <div
            role="alert"
            className="bg-red-50 border border-red-200 text-red-800 rounded-xl px-5 py-3.5 text-sm flex items-center justify-between gap-4"
          >
            <div>
              <span className="font-semibold">Lỗi xử lý: </span>
              <span>{errorMessage}</span>
            </div>
            <button
              type="button"
              onClick={() => setErrorMessage(null)}
              className="text-xs font-medium text-red-700 hover:underline whitespace-nowrap shrink-0 cursor-pointer"
            >
              Đóng
            </button>
          </div>
        )}

        {statusMessage && !errorMessage && (
          <div className="bg-emerald-50 border border-emerald-200 text-emerald-900 rounded-xl px-5 py-3 text-sm flex items-center justify-between gap-4">
            <span>{statusMessage}</span>
            <button
              type="button"
              onClick={() => setStatusMessage(null)}
              className="text-xs font-medium text-emerald-800 hover:underline whitespace-nowrap shrink-0 cursor-pointer"
            >
              Ẩn
            </button>
          </div>
        )}

        {/* TAB 1: PROCESSOR */}
        {activeTab === 'processor' && (
          <div className="space-y-6">
            {/* Hero / Upload & Sample Controls */}
            <section className="bg-white border border-slate-200 rounded-xl p-6">
              <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6">
                <div className="space-y-1.5 max-w-2xl">
                  <h1 className="font-display text-2xl font-bold text-slate-900 text-balance">
                    Bóc tách Thời khóa biểu tổng quát theo từng môn học
                  </h1>
                  <p className="text-sm text-slate-600 leading-relaxed">
                    Quét động toàn bộ các khối tuần liên tiếp (15, 17, 21 tuần hoặc bất kỳ), tự động
                    bóc tách tiêu đề bằng Regex, phân công giảng viên và xuất từng file{' '}
                    <span className="font-mono text-slate-800">&lt;Class&gt;_&lt;Subject&gt;.xlsx</span>{' '}
                    chuẩn hóa.
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-3 shrink-0">
                  <button
                    type="button"
                    onClick={() => handleLoadSample('17w')}
                    disabled={isLoading}
                    className="px-3.5 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-50"
                  >
                    Thử mẫu 17 tuần (PC2204)
                  </button>
                  <button
                    type="button"
                    onClick={() => handleLoadSample('21w')}
                    disabled={isLoading}
                    className="px-3.5 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-50"
                  >
                    Thử mẫu 21 tuần (SE2102)
                  </button>
                  <button
                    type="button"
                    onClick={() => setActiveTab('python')}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Code2 className="w-4 h-4" />
                    Lấy mã nguồn Python Tkinter
                  </button>
                </div>
              </div>

              {/* Dropzone */}
              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDragging(true);
                }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsDragging(false);
                  handleFileSelect(e.dataTransfer.files?.[0]);
                }}
                onClick={() => fileInputRef.current?.click()}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    fileInputRef.current?.click();
                  }
                }}
                className={`mt-5 border border-dashed rounded-lg p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4 transition-colors cursor-pointer ${
                  isDragging
                    ? 'border-blue-600 bg-blue-50/50'
                    : 'border-slate-300 hover:border-slate-400 bg-slate-50/60'
                }`}
              >
                <div className="flex items-center gap-3.5 min-w-0">
                  <div className="w-10 h-10 rounded-lg bg-white border border-slate-200 flex items-center justify-center shrink-0 text-blue-700">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div className="min-w-0">
                    <p className="text-sm font-medium text-slate-900 truncate">
                      {fileName ? `Đang mở: ${fileName}` : 'Kéo thả file Schedule.xlsx vào đây hoặc bấm để chọn file'}
                    </p>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Hỗ trợ cấu trúc bảng tính nhiều khối tuần (Tuần 1-3, Tuần 4-6... đến hết sheet)
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2 shrink-0" onClick={(e) => e.stopPropagation()}>
                  <button
                    type="button"
                    onClick={() => handleDownloadSampleExcel('17w')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Tải file mẫu 17 tuần
                  </button>
                  <button
                    type="button"
                    onClick={() => handleDownloadSampleExcel('21w')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Tải file mẫu 21 tuần
                  </button>
                </div>
              </div>
            </section>

            {/* Empty State */}
            {!parseResult && !isLoading && (
              <section className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-3">
                <p className="text-base font-semibold text-slate-900">
                  Chưa có dữ liệu thời khóa biểu nào được nạp
                </p>
                <p className="text-sm text-slate-600 max-w-lg mx-auto">
                  Hãy chọn file <span className="font-mono text-slate-800">Schedule.xlsx</span> từ máy
                  của bạn hoặc nhấn <span className="font-medium text-slate-900">Thử mẫu 17 tuần</span>{' '}
                  phía trên để kiểm tra trực tiếp cơ chế quét động và xuất file Excel.
                </p>
              </section>
            )}

            {/* Loading Skeleton */}
            {isLoading && (
              <section className="bg-white border border-slate-200 rounded-xl p-6 space-y-4">
                <div className="h-5 w-48 bg-slate-200 rounded animate-pulse" />
                <div className="grid grid-cols-2 md:grid-cols-6 gap-4">
                  {Array.from({ length: 6 }).map((_, i) => (
                    <div key={i} className="h-12 bg-slate-100 rounded animate-pulse" />
                  ))}
                </div>
                <div className="h-44 bg-slate-100 rounded animate-pulse" />
              </section>
            )}

            {/* Populated State */}
            {parseResult && !isLoading && (
              <>
                {/* Extracted Metadata Grid */}
                <section className="bg-white border border-slate-200 rounded-xl p-6 space-y-5">
                  <div className="flex flex-wrap items-center justify-between gap-4">
                    <h2 className="font-display text-base font-bold text-slate-900">
                      Thông tin tự động bóc tách từ tiêu đề &amp; cấu trúc bảng
                    </h2>
                    <div className="text-xs text-slate-500 flex items-center gap-2 tabular-nums">
                      <span>{parseResult.blocks.length} khối tuần</span>
                      <span aria-hidden="true">·</span>
                      <span>{parseResult.records.length} buổi học hợp lệ</span>
                      {parseResult.ignoredNotes.length > 0 && (
                        <>
                          <span aria-hidden="true">·</span>
                          <button
                            type="button"
                            onClick={() => setShowIgnoredNotes((v) => !v)}
                            className="text-amber-700 hover:underline font-medium cursor-pointer"
                          >
                            Đã lọc {parseResult.ignoredNotes.length} ô ghi chú/ngày nghỉ
                          </button>
                        </>
                      )}
                    </div>
                  </div>

                  <dl className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-x-6 gap-y-4 pt-2 border-t border-slate-100">
                    <div>
                      <dt className="text-xs text-slate-500">Lớp học (Class)</dt>
                      <dd className="mt-1 font-mono text-sm font-semibold text-slate-900">
                        {parseResult.metadata.className}
                      </dd>
                    </div>
                    <div>
                      <dt className="text-xs text-slate-500">Phòng học (Room)</dt>
                      <dd className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                        {parseResult.metadata.room}
                      </dd>
                    </div>
                    <div>
                      <dt className="text-xs text-slate-500">Học kỳ (Term)</dt>
                      <dd className="mt-1 font-mono text-sm font-semibold text-slate-900">
                        {parseResult.metadata.term}
                      </dd>
                    </div>
                    <div>
                      <dt className="text-xs text-slate-500">Khoảng ngày</dt>
                      <dd className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                        {parseResult.metadata.dateBeginStr} – {parseResult.metadata.dateEndStr}
                      </dd>
                    </div>
                    <div>
                      <dt className="text-xs text-slate-500">Ca học nhận diện</dt>
                      <dd className="mt-1 text-sm font-semibold text-slate-900">
                        {parseResult.metadata.shiftLabel}
                      </dd>
                    </div>
                    <div>
                      <dt className="text-xs text-slate-500">Tổng thời lượng</dt>
                      <dd className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                        {parseResult.metadata.totalWeeks} tuần ({parseResult.metadata.totalWeeks * 7} ngày)
                      </dd>
                    </div>
                  </dl>

                  {/* Dynamic Block Scan Trail */}
                  <div className="pt-4 border-t border-slate-100">
                    <p className="text-xs font-medium text-slate-500 mb-2">
                      Các khối tuần đã quét động từ trên xuống dưới:
                    </p>
                    <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 text-xs text-slate-700 font-mono tabular-nums">
                      {parseResult.blocks.map((blk, i) => (
                        <span key={`${blk.label}-${i}`}>
                          <span className="font-semibold text-slate-900">{blk.label}</span> (
                          {blk.weeksCount}T · {blk.slotsRange} · {blk.sessionsFound} buổi)
                        </span>
                      ))}
                    </div>
                  </div>

                  {showIgnoredNotes && parseResult.ignoredNotes.length > 0 && (
                    <div className="pt-4 border-t border-slate-100">
                      <p className="text-xs font-medium text-amber-800 mb-2">
                        Danh sách ô ghi chú sự kiện / ngày nghỉ đã tự động bỏ qua:
                      </p>
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs text-left">
                          <thead>
                            <tr className="border-b border-slate-200 text-slate-500">
                              <th className="py-1.5 pr-4 font-medium">Nội dung ô</th>
                              <th className="py-1.5 pr-4 font-medium">Ngày</th>
                              <th className="py-1.5 font-medium">Tiết</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-mono tabular-nums">
                            {parseResult.ignoredNotes.map((note, idx) => (
                              <tr key={idx}>
                                <td className="py-1.5 pr-4 font-sans text-slate-800">{note.text}</td>
                                <td className="py-1.5 pr-4 text-slate-600">{note.dateStr}</td>
                                <td className="py-1.5 text-slate-600">Tiết {note.slot}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </section>

                {/* Subject & Lecturer Assignment Table */}
                <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                  <div className="px-6 py-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h2 className="font-display text-base font-bold text-slate-900">
                        Phân công Giảng viên &amp; Xuất file từng môn học
                      </h2>
                      <p className="text-xs text-slate-500 mt-0.5">
                        Nhập tên Giảng viên cho mỗi môn học tìm thấy, sau đó xuất từng file hoặc tải toàn bộ dạng ZIP.
                      </p>
                    </div>

                    <button
                      type="button"
                      onClick={handleExportAllZip}
                      disabled={isExporting}
                      className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-50"
                    >
                      <FolderArchive className="w-4 h-4" />
                      {isExporting ? 'Đang tạo file...' : `Xuất tất cả ${parseResult.subjects.length} file (.zip)`}
                    </button>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-left">
                      <thead className="bg-slate-50 border-b border-slate-200 text-xs text-slate-600">
                        <tr>
                          <th className="py-3 px-6 font-semibold">Mã môn (Subject)</th>
                          <th className="py-3 px-4 font-semibold text-right">Số buổi</th>
                          <th className="py-3 px-4 font-semibold">Buổi đầu</th>
                          <th className="py-3 px-4 font-semibold">Buổi cuối</th>
                          <th className="py-3 px-4 font-semibold">Tên Giảng viên (Lecturer)</th>
                          <th className="py-3 px-6 font-semibold text-right">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {subjectSummaries.map((item) => (
                          <tr key={item.subject} className="hover:bg-slate-50/70 transition-colors">
                            <td className="py-2.5 px-6 font-mono font-semibold text-slate-900">
                              {item.subject}
                            </td>
                            <td className="py-2.5 px-4 font-mono tabular-nums text-right text-slate-700">
                              {item.count}
                            </td>
                            <td className="py-2.5 px-4 font-mono tabular-nums text-slate-600">
                              {item.firstDate}
                            </td>
                            <td className="py-2.5 px-4 font-mono tabular-nums text-slate-600">
                              {item.lastDate}
                            </td>
                            <td className="py-2 px-4 min-w-[240px]">
                              <input
                                type="text"
                                value={item.lecturer}
                                placeholder="Nhập tên giảng viên..."
                                onChange={(e) =>
                                  setLecturerMap((prev) => ({
                                    ...prev,
                                    [item.subject]: e.target.value,
                                  }))
                                }
                                className="w-full px-3 py-1.5 text-sm bg-white border border-slate-300 rounded-md focus:outline-none focus:border-blue-600 focus:ring-1 focus:ring-blue-600"
                              />
                            </td>
                            <td className="py-2 px-6 text-right">
                              <div className="inline-flex items-center justify-end gap-3">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setSelectedSubject(item.subject);
                                    setActiveTab('preview');
                                  }}
                                  className="text-xs font-medium text-slate-600 hover:text-slate-900 hover:underline whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  Xem lịch
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleExportSingleSubject(item.subject)}
                                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  <Download className="w-3.5 h-3.5" />
                                  {parseResult.metadata.className}_{item.subject}.xlsx
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              </>
            )}
          </div>
        )}

        {/* TAB 2: PREVIEW */}
        {activeTab === 'preview' && (
          <div className="space-y-6">
            {!parseResult ? (
              <section className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-4">
                <p className="text-base font-semibold text-slate-900">
                  Chưa có dữ liệu để xem trước
                </p>
                <p className="text-sm text-slate-600">
                  Hãy chọn file Schedule.xlsx hoặc nạp dữ liệu mẫu 17 tuần / 21 tuần để kiểm tra 8 cột dữ liệu chuẩn.
                </p>
                <div className="flex justify-center gap-3">
                  <button
                    type="button"
                    onClick={() => handleLoadSample('17w')}
                    className="px-4 py-2 text-sm font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg cursor-pointer"
                  >
                    Nạp mẫu 17 tuần (PC2204)
                  </button>
                </div>
              </section>
            ) : (
              <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-xs font-medium text-slate-500 mr-1">Chọn môn học:</span>
                    <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg">
                      {parseResult.subjects.map((subj) => (
                        <button
                          key={subj}
                          type="button"
                          onClick={() => setSelectedSubject(subj)}
                          className={`px-3 py-1 text-xs font-mono font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                            selectedSubject === subj
                              ? 'bg-white text-slate-900 shadow-sm'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          {subj}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="flex items-center gap-3">
                    <div className="relative">
                      <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={previewSearch}
                        onChange={(e) => setPreviewSearch(e.target.value)}
                        placeholder="Lọc ngày, tiết, buổi..."
                        className="pl-9 pr-3 py-1.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-blue-600"
                      />
                    </div>
                    <button
                      type="button"
                      onClick={() => handleExportSingleSubject(selectedSubject)}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-sm font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <Download className="w-4 h-4" />
                      Xuất {parseResult.metadata.className}_{selectedSubject}.xlsx
                    </button>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-sm text-center">
                    <thead className="bg-[#DCE6F1] border-b border-slate-300 text-xs font-bold text-slate-900">
                      <tr>
                        <th className="py-3 px-4">Term</th>
                        <th className="py-3 px-4">Date</th>
                        <th className="py-3 px-4">Slot</th>
                        <th className="py-3 px-4">Room</th>
                        <th className="py-3 px-4">Lecturer</th>
                        <th className="py-3 px-4">Subject</th>
                        <th className="py-3 px-4">Session No.</th>
                        <th className="py-3 px-4">Class</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 tabular-nums">
                      {previewSessions.map((row) => (
                        <tr
                          key={`${row.subject}-${row.sessionNo}`}
                          className="even:bg-slate-50/70 hover:bg-blue-50/30"
                        >
                          <td className="py-2.5 px-4 text-slate-700">{row.term}</td>
                          <td className="py-2.5 px-4 font-mono text-slate-900">{row.dateStr}</td>
                          <td className="py-2.5 px-4 font-mono text-slate-800">{row.slot}</td>
                          <td className="py-2.5 px-4 font-mono text-slate-700">{row.room}</td>
                          <td className="py-2.5 px-4 text-slate-800">
                            {row.lecturer || <span className="text-slate-400 italic">Chưa gán</span>}
                          </td>
                          <td className="py-2.5 px-4 font-mono font-semibold text-slate-900">
                            {row.subject}
                          </td>
                          <td className="py-2.5 px-4 font-mono font-semibold text-blue-700">
                            {row.sessionNo}
                          </td>
                          <td className="py-2.5 px-4 font-mono text-slate-800">{row.className}</td>
                        </tr>
                      ))}
                      {previewSessions.length === 0 && (
                        <tr>
                          <td colSpan={8} className="py-8 text-slate-500">
                            Không tìm thấy buổi học nào khớp với bộ lọc.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>
            )}
          </div>
        )}

        {/* TAB 3: PYTHON SOURCE */}
        {activeTab === 'python' && (
          <div className="space-y-6">
            <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                <div className="space-y-1">
                  <h2 className="font-display text-base font-bold text-slate-900">
                    Mã nguồn ứng dụng Python Desktop (Tkinter + OpenPyXL)
                  </h2>
                  <p className="text-xs text-slate-600">
                    {PYTHON_FILES[activePyFile].description}
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
                    {(Object.keys(PYTHON_FILES) as PythonFileKey[]).map((key) => (
                      <button
                        key={key}
                        type="button"
                        onClick={() => setActivePyFile(key)}
                        className={`px-3 py-1.5 text-xs font-mono font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          activePyFile === key
                            ? 'bg-white text-slate-900 shadow-sm'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        {key}
                      </button>
                    ))}
                  </div>

                  <button
                    type="button"
                    onClick={() => handleCopyCode(activePyFile)}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    {copiedKey === activePyFile ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-emerald-700" />
                        Đã sao chép
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        Sao chép mã
                      </>
                    )}
                  </button>

                  <button
                    type="button"
                    onClick={() => handleDownloadPythonFile(activePyFile)}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Tải {activePyFile}
                  </button>
                </div>
              </div>

              {/* Run command hint */}
              <div className="px-6 py-2.5 bg-slate-900 text-slate-300 text-xs font-mono border-b border-slate-800 flex flex-wrap items-center gap-x-6 gap-y-1">
                <span>Cài thư viện: pip install openpyxl</span>
                <span>Chạy GUI: python schedule_splitter.py</span>
                <span>Tạo file mẫu 17 &amp; 21 tuần: python generate_sample_schedule.py</span>
              </div>

              <pre className="p-6 bg-slate-950 text-slate-100 text-xs leading-relaxed overflow-x-auto max-h-[620px] font-mono">
                <code>{PYTHON_FILES[activePyFile].content}</code>
              </pre>
            </section>
          </div>
        )}

        {/* TAB 4: ARCHITECTURE */}
        {activeTab === 'architecture' && (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <section className="bg-white border border-slate-200 rounded-xl p-6 space-y-4">
              <h2 className="font-display text-base font-bold text-slate-900">
                01. So sánh Macro VBA cũ và Cơ chế Quét Động mới
              </h2>
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left">
                  <thead className="border-b border-slate-200 text-xs text-slate-500">
                    <tr>
                      <th className="py-2 pr-4 font-medium">Tiêu chí</th>
                      <th className="py-2 pr-4 font-medium">VBA cũ (Hardcode)</th>
                      <th className="py-2 font-medium">Python Dynamic Scanning</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-xs">
                    <tr>
                      <td className="py-2.5 pr-4 font-medium text-slate-900">Số lượng tuần</td>
                      <td className="py-2.5 pr-4 text-red-700">Cố định 5 khối (15 tuần)</td>
                      <td className="py-2.5 text-emerald-700 font-medium">
                        Bất kỳ (15, 17, 21+ tuần)
                      </td>
                    </tr>
                    <tr>
                      <td className="py-2.5 pr-4 font-medium text-slate-900">Thông tin Lớp / Kỳ</td>
                      <td className="py-2.5 pr-4 text-slate-600">Nhập tay dễ sai sót</td>
                      <td className="py-2.5 text-slate-900">Regex tự động từ tiêu đề A1</td>
                    </tr>
                    <tr>
                      <td className="py-2.5 pr-4 font-medium text-slate-900">Nhận diện Ca học</td>
                      <td className="py-2.5 pr-4 text-slate-600">Cố định chỉ số dòng</td>
                      <td className="py-2.5 text-slate-900">Đọc nhãn Tiết 1–5 / Tiết 6–10</td>
                    </tr>
                    <tr>
                      <td className="py-2.5 pr-4 font-medium text-slate-900">Ghi chú ngày nghỉ</td>
                      <td className="py-2.5 pr-4 text-slate-600">Dễ tách nhầm thành môn</td>
                      <td className="py-2.5 text-slate-900">Lọc tự động tiếng Việt &amp; từ khóa</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>

            <section className="bg-white border border-slate-200 rounded-xl p-6 space-y-4">
              <h2 className="font-display text-base font-bold text-slate-900">
                02. Quy trình thuật toán xử lý 5 bước
              </h2>
              <ol className="space-y-3 text-sm text-slate-700">
                <li>
                  <span className="font-mono font-semibold text-slate-900 mr-2">01.</span>
                  Bóc tách tiêu đề <span className="font-mono text-xs text-slate-900">Term</span>,{' '}
                  <span className="font-mono text-xs text-slate-900">Date_Begin</span>,{' '}
                  <span className="font-mono text-xs text-slate-900">Date_End</span>,{' '}
                  <span className="font-mono text-xs text-slate-900">Class</span>,{' '}
                  <span className="font-mono text-xs text-slate-900">Room</span> bằng Regex.
                </li>
                <li>
                  <span className="font-mono font-semibold text-slate-900 mr-2">02.</span>
                  Duyệt dòng từ <span className="font-mono text-xs">r = 1</span> tới{' '}
                  <span className="font-mono text-xs">max_row</span>, bắt mọi khối bắt đầu bằng{' '}
                  <span className="font-mono text-xs">Tuần X-Y</span> và xác định số cột ngày động (7, 14, 21 ngày...).
                </li>
                <li>
                  <span className="font-mono font-semibold text-slate-900 mr-2">03.</span>
                  Tăng <span className="font-mono text-xs">current_date += timedelta(days=1)</span> liên tục
                  qua mọi cột ngày để khớp chính xác ngày học thực tế (17 tuần = 119 ngày, 21 tuần = 147 ngày).
                </li>
                <li>
                  <span className="font-mono font-semibold text-slate-900 mr-2">04.</span>
                  Chuẩn hóa mã môn in hoa, bỏ qua ghi chú sự kiện như{' '}
                  <span className="italic">&ldquo;Nghỉ lễ 2/9&rdquo;</span>,{' '}
                  <span className="italic">&ldquo;Khai giảng K22&rdquo;</span>.
                </li>
                <li>
                  <span className="font-mono font-semibold text-slate-900 mr-2">05.</span>
                  Sắp xếp theo trình tự thời gian tăng dần, đánh{' '}
                  <span className="font-mono text-xs">Session No. = 1..N</span> và xuất file{' '}
                  <span className="font-mono text-xs">&lt;Class&gt;_&lt;Subject&gt;.xlsx</span> 8 cột chuẩn.
                </li>
              </ol>
            </section>
          </div>
        )}
      </main>
    </div>
  );
}
