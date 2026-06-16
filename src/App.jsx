import { useState, useEffect, useMemo, useCallback } from 'react';
import { PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { CheckCircle, AlertCircle, BookOpen, Users, Calendar, Settings, ArrowRight, LayoutDashboard, Clock, MapPin, Sparkles, Save, Edit3, Lock, User, LogOut, Award, Download } from 'lucide-react';
import * as XLSX from 'xlsx';

const normalizeHeader = (header) => String(header ?? '').trim().toLowerCase().replace(/\s+/g, ' ');

const getRowValue = (row, keys) => {
  const wantedHeaders = new Set(keys.map(normalizeHeader));

  for (const [header, value] of Object.entries(row || {})) {
    if (wantedHeaders.has(normalizeHeader(header)) && value !== undefined && value !== null && String(value).trim()) {
      return String(value).trim();
    }
  }

  return '';
};

const createSheetFieldAliases = (value, keys) => {
  return keys.reduce((aliases, key) => {
    aliases[key] = value;
    aliases[` ${key}`] = value;
    aliases[`${key} `] = value;
    return aliases;
  }, {});
};

const normalizeRow = (row) => {
  return {
    name: getRowValue(row, ['Name', 'name', 'Student Name', 'StudentName', 'ชื่อนักเรียน', 'ชื่อ', 'ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อ - สกุล']),
    subject: getRowValue(row, ['Subject', 'subject', 'วิชา', 'รายวิชา']) || '-',
    date: getRowValue(row, ['Date', 'date', 'วันที่', 'วันที่เช็คชื่อ']) || '-',
    status: getRowValue(row, ['Status', 'status', 'สถานะ']) || '-',
    assignment: getRowValue(row, ['Assignment', 'assignment', 'ชื่องาน', 'งาน']) || '-',
    dueDate: getRowValue(row, ['DueDate', 'dueDate', 'Due Date', 'กำหนดส่ง', 'วันส่ง']) || '-',
    studentId: getRowValue(row, ['StudentID', 'studentId', 'Student ID', 'เลขที่', 'รหัส', 'รหัสนักเรียน']) || '-',
    className: getRowValue(row, ['Class', 'className', 'ClassName', 'ชั้นเรียน', 'ชั้น', 'ห้อง']) || '-'
  };
};

const createAttendancePayloadRow = ({ student, subject, date, status, className }) => ({
  ...createSheetFieldAliases(student.name, ['Name', 'name', 'Student Name', 'StudentName', 'ชื่อนักเรียน', 'ชื่อ', 'ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อ - สกุล']),
  ...createSheetFieldAliases(student.studentId, ['StudentID', 'studentId', 'Student ID', 'เลขที่', 'รหัสนักเรียน']),
  ...createSheetFieldAliases(className, ['Class', 'className', 'ClassName', 'ชั้นเรียน', 'ชั้น', 'ห้อง']),
  ...createSheetFieldAliases(subject, ['Subject', 'subject', 'วิชา', 'รายวิชา']),
  ...createSheetFieldAliases(date, ['Date', 'date', 'วันที่', 'วันที่เช็คชื่อ']),
  ...createSheetFieldAliases(status, ['Status', 'status', 'สถานะ'])
});

const createAssignmentPayloadRow = ({ student, subject, assignment, dueDate, status, className }) => ({
  ...createSheetFieldAliases(student.name, ['Name', 'name', 'Student Name', 'StudentName', 'ชื่อนักเรียน', 'ชื่อ', 'ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อ - สกุล']),
  ...createSheetFieldAliases(student.studentId, ['StudentID', 'studentId', 'Student ID', 'เลขที่', 'รหัสนักเรียน']),
  ...createSheetFieldAliases(className, ['Class', 'className', 'ClassName', 'ชั้นเรียน', 'ชั้น', 'ห้อง']),
  ...createSheetFieldAliases(subject, ['Subject', 'subject', 'วิชา', 'รายวิชา']),
  ...createSheetFieldAliases(assignment, ['Assignment', 'assignment', 'ชื่องาน', 'งาน']),
  ...createSheetFieldAliases(dueDate, ['DueDate', 'dueDate', 'Due Date', 'กำหนดส่ง', 'วันส่ง']),
  ...createSheetFieldAliases(status, ['Status', 'status', 'สถานะ'])
});


const getMonday = (d) => {
  const date = new Date(d);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is sunday
  return new Date(date.setDate(diff));
};

const toLocalDateString = (d) => {
  const date = new Date(d);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};



const thaiMonthName = (monthIdx) => {
  const months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  return months[monthIdx] || '';
};


const formatHeaderDate = (dateStr) => {
  if (!dateStr || dateStr === '-') return '';
  const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const [, m, d] = dateStr.split('-').map(Number); // local time — ไม่ใช้ new Date(str) ที่ parse เป็น UTC
  return `${d} ${months[m - 1]}`;
};

const splitName = (fullName) => {
  if (!fullName) return { firstName: '-', lastName: '-' };
  const trimmed = fullName.trim();
  const parts = trimmed.split(/\s+/);
  const firstName = parts[0] || '-';
  const lastName = parts.slice(1).join(' ') || '-';
  return { firstName, lastName };
};

// ===== Cross-Class Subject Configuration =====
// Classes automatically included for ลูกเสือ
const SCOUT_CLASSES = ['ป.4', 'ป.5', 'ป.6', 'ม.1', 'ม.2', 'ม.3'];

// Keywords in subject name that indicate a cross-class (all-class) subject
const CROSS_CLASS_KEYWORDS = ['ชุมนุม', 'ลูกเสือ'];

const isCrossClassSubject = (subjectName) => {
  if (!subjectName) return false;
  return CROSS_CLASS_KEYWORDS.some(kw => subjectName.includes(kw));
};

const isScoutSubject = (subjectName) => {
  if (!subjectName) return false;
  return subjectName.includes('ลูกเสือ');
};

const isClubSubject = (subjectName) => {
  if (!subjectName) return false;
  return subjectName.includes('ชุมนุม');
};

const DecorativeBackground = () => (
  <div className="fixed inset-0 overflow-hidden pointer-events-none z-[-1] bg-slate-50">
    <div className="absolute top-[-10%] left-[-10%] w-96 h-96 bg-purple-300 rounded-full mix-blend-multiply filter blur-[80px] opacity-60 animate-blob"></div>
    <div className="absolute top-[20%] right-[-10%] w-[30rem] h-[30rem] bg-cyan-300 rounded-full mix-blend-multiply filter blur-[100px] opacity-60 animate-blob animation-delay-2000"></div>
    <div className="absolute bottom-[-20%] left-[20%] w-[40rem] h-[40rem] bg-pink-200 rounded-full mix-blend-multiply filter blur-[120px] opacity-60 animate-blob animation-delay-4000"></div>
  </div>
);

// Timetable Data and Fuzzy Matching Constants
const TIMETABLE_DAYS = [
  { name: 'จันทร์', color: 'bg-yellow-400 text-slate-800 font-bold', code: 'Mon', index: 0, hoverClass: 'hover:bg-yellow-50/50' },
  { name: 'อังคาร', color: 'bg-pink-400 text-white font-bold', code: 'Tue', index: 1, hoverClass: 'hover:bg-pink-50/50' },
  { name: 'พุธ', color: 'bg-emerald-500 text-white font-bold', code: 'Wed', index: 2, hoverClass: 'hover:bg-emerald-50/50' },
  { name: 'พฤหัสบดี', color: 'bg-orange-500 text-white font-bold', code: 'Thu', index: 3, hoverClass: 'hover:bg-orange-50/50' },
  { name: 'ศุกร์', color: 'bg-sky-500 text-white font-bold', code: 'Fri', index: 4, hoverClass: 'hover:bg-sky-50/50' },
];

const TIMETABLE_PERIODS = [
  { number: 1, time: '8.30 - 9.30' },
  { number: 2, time: '9.30 - 10.30' },
  { number: 3, time: '10.30 - 11.30' },
  { number: 'break', label: 'พักกลางวัน', time: '11.30 - 12.30' },
  { number: 4, time: '12.30 - 13.30' },
  { number: 5, time: '13.30 - 14.30' },
  { number: 6, time: '14.30 - 15.30' },
  { number: 7, time: '-' }
];

const SCHEDULE = {
  'จันทร์': {
    2: { subject: 'คณิตฯเพิ่มเติม', class: 'ม.1' },
    4: { subject: 'ดนตรี', class: 'ป.6' },
    5: { subject: 'ซ่อมเสริม(ดนตรี)', class: 'ป.6' },
    6: { subject: 'ศิลปะ', class: 'ป.5' },
  },
  'อังคาร': {
    2: { subject: 'ศิลปะ', class: 'ม.1' },
    3: { subject: 'ดนตรี', class: 'ม.1' },
    4: { subject: 'ดนตรี', class: 'ป.4' },
    5: { subject: 'ซ่อมเสริม(ดนตรี)', class: 'ป.4' },
  },
  'พุธ': {
    1: { subject: 'ศิลปะ', class: 'ป.4' },
    2: { subject: 'ศิลปะ', class: 'ม.3' },
    3: { subject: 'ดนตรี', class: 'ม.3' },
    4: { subject: 'ดนตรี', class: 'ป.5' },
    5: { subject: 'ซ่อมเสริม(ดนตรี)', class: 'ป.5' },
    6: { subject: 'ชุมนุม1', class: 'all' },
  },
  'พฤหัสบดี': {
    2: { subject: 'ศิลปะ', class: 'ม.2' },
    3: { subject: 'ดนตรี', class: 'ม.2' },
    4: { subject: 'วิทย์คำนวณ', class: 'ม.1' },
    5: { subject: 'วิทย์คำนวณ', class: 'ม.3' },
    6: { subject: 'ลูกเสือ', class: 'all' },
  },
  'ศุกร์': {
    2: { subject: 'ศิลปะ', class: 'ป.6' },
    3: { subject: 'คณิตฯเพิ่มเติม', class: 'ม.1' },
    4: { subject: 'กิจกรรมอิสระ4', class: 'ป.1' },
    5: { subject: 'วิทย์คำนวณ', class: 'ม.2' },
  }
};

const cleanKey = (str) => {
  if (!str) return '';
  return str
    .replace(/[\s()-]/g, '')
    .replace(/วิทยาการ/g, 'วิทย')
    .replace(/วิทยา/g, 'วิทย')
    .replace(/วิทย์/g, 'วิทย')
    .replace(/คณิตศาสตร์/g, 'คณิต')
    .replace(/คณิตฯ/g, 'คณิต')
    .toLowerCase();
};

const resolveSubjectFromTimetable = (timetableSubject, timetableClass, subjectList) => {
  if (!subjectList || subjectList.length === 0) return '';
  
  const cleanedSub = timetableSubject.trim();
  const cleanedClass = timetableClass.trim();
  
  // 1. Exact match
  const exactMatch = subjectList.find(s => s === cleanedSub || s === `${cleanedSub} ${cleanedClass}`);
  if (exactMatch) return exactMatch;
  
  // 2. Special case: Art / Music (ศิลปะ / ดนตรี)
  const isArtOrMusic = (cleanedSub.includes('ศิลปะ') || cleanedSub.includes('ดนตรี')) && !cleanedSub.includes('ซ่อมเสริม');
  if (isArtOrMusic) {
    const artMusicMatch = subjectList.find(s => {
      const lowerS = s.toLowerCase();
      const hasArt = lowerS.includes('ศิลปะ');
      const hasMusic = lowerS.includes('ดนตรี');
      const hasClass = cleanedClass === 'all' || lowerS.includes(cleanedClass.toLowerCase());
      return hasArt && hasMusic && hasClass;
    });
    if (artMusicMatch) return artMusicMatch;
    
    const singleMatch = subjectList.find(s => {
      const lowerS = s.toLowerCase();
      const hasArtOrMusic = lowerS.includes('ศิลปะ') || lowerS.includes('ดนตรี');
      const hasClass = cleanedClass === 'all' || lowerS.includes(cleanedClass.toLowerCase());
      return hasArtOrMusic && hasClass;
    });
    if (singleMatch) return singleMatch;
  }
  
  // 3. General fuzzy matcher
  const targetKey = cleanKey(cleanedSub);
  
  let match = subjectList.find(s => {
    const sClean = cleanKey(s);
    const hasClass = cleanedClass === 'all' || sClean.includes(cleanKey(cleanedClass));
    return hasClass && (sClean.includes(targetKey) || targetKey.includes(sClean));
  });
  
  if (match) return match;
  
  match = subjectList.find(s => {
    const sClean = cleanKey(s);
    return sClean.includes(targetKey) || targetKey.includes(sClean);
  });
  
  if (match) return match;

  if (targetKey.length >= 3) {
    const shortKey = targetKey.substring(0, 4);
    match = subjectList.find(s => {
      const sClean = cleanKey(s);
      return sClean.includes(shortKey);
    });
    if (match) return match;
  }
  
  return '';
};

const TimetableGrid = ({ subjectList, onCellClick }) => {
  return (
    <div className="glass-card rounded-[2rem] p-6 shadow-xl border border-slate-100/50 bg-white/40 animate-fade-in-up">
      {/* Header */}
      <div className="text-center mb-6">
        <span className="bg-indigo-50 text-indigo-700 text-xs font-bold px-3 py-1 rounded-full uppercase tracking-wider">ตารางเรียนครูผู้สอน</span>
        <h2 className="text-2xl font-black text-slate-800 mt-2 flex items-center justify-center gap-2">
          <Sparkles className="text-yellow-500 fill-yellow-400" size={24} /> โรงเรียนบ้านปากยาง
        </h2>
        <div className="flex flex-wrap items-center justify-center gap-x-6 gap-y-2 mt-3 text-sm text-slate-500 font-medium bg-white/40 backdrop-blur-md py-2 px-6 rounded-2xl border border-slate-200/50 inline-flex">
          <span>ครูผู้สอน: <strong className="text-indigo-600">ครูวีรวัฒน์</strong></span>
          <span className="w-1.5 h-1.5 bg-slate-300 rounded-full"></span>
          <span>จำนวนภาระงาน: <strong className="text-slate-700">23 คาบ</strong></span>
          <span className="w-1.5 h-1.5 bg-slate-300 rounded-full"></span>
          <span>ปีการศึกษา: <strong className="text-slate-700">2569</strong></span>
        </div>
      </div>
      
      {/* Table Container */}
      <div className="overflow-x-auto rounded-2xl border border-slate-200/60 bg-white/50 backdrop-blur-md">
        <table className="w-full min-w-[950px] border-collapse text-center">
          <thead>
            {/* Period numbers */}
            <tr className="bg-slate-50/80 text-xs font-bold text-slate-500 border-b border-slate-200">
              <th rowSpan={2} className="p-4 border-r border-slate-200 w-[110px] bg-slate-100/70 font-black text-slate-700">วัน</th>
              {TIMETABLE_PERIODS.map((p, idx) => {
                if (p.number === 'break') {
                  return (
                    <th key={idx} rowSpan={2} className="p-2 border-r border-slate-200 w-[50px] bg-amber-50/40 text-amber-800 font-black">
                      <div className="writing-vertical-lr select-none tracking-widest uppercase text-xs mx-auto">
                        พักกลางวัน
                      </div>
                    </th>
                  );
                }
                return (
                  <th key={idx} className="p-2 border-r border-slate-200 text-sm font-black text-slate-700 bg-slate-100/40">
                    คาบที่ {p.number}
                  </th>
                );
              })}
            </tr>
            {/* Period timings */}
            <tr className="bg-slate-50/50 text-[10px] font-semibold text-slate-400 border-b border-slate-200">
              {TIMETABLE_PERIODS.filter(p => p.number !== 'break').map((p, idx) => (
                <th key={idx} className="p-2 border-r border-slate-200 font-medium">
                  {p.time}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {TIMETABLE_DAYS.map((day, dayIdx) => {
              const rowSchedule = SCHEDULE[day.name] || {};
              return (
                <tr key={dayIdx} className="border-b border-slate-200/60 hover:bg-slate-50/20 transition-all">
                  {/* Day Label Cell */}
                  <td className={`p-4 border-r border-slate-200 ${day.color} text-sm font-extrabold text-center select-none shadow-sm`}>
                    {day.name}
                  </td>
                  
                  {/* Periods Cells */}
                  {TIMETABLE_PERIODS.map((period, pIdx) => {
                    if (period.number === 'break') {
                      return (
                        <td key={pIdx} className="p-2 border-r border-slate-200 bg-amber-50/10 text-amber-800/30 text-xs font-bold select-none">
                          
                        </td>
                      );
                    }
                    
                    const slot = rowSchedule[period.number];
                    if (!slot) {
                      return (
                        <td key={pIdx} className="p-2 border-r border-slate-200 bg-slate-50/10 text-slate-300 text-xs italic select-none">
                          -
                        </td>
                      );
                    }
                    
                    const resolvedSub = resolveSubjectFromTimetable(slot.subject, slot.class, subjectList);
                    const isMatched = !!resolvedSub;
                    
                    return (
                      <td key={pIdx} className="p-2 border-r border-slate-200">
                        <button
                          type="button"
                          onClick={() => onCellClick(day.name, slot)}
                          className={`w-full h-full min-h-[76px] p-2.5 rounded-2xl flex flex-col justify-between items-center text-center cursor-pointer transition-all transform hover:scale-[1.02] active:scale-[0.98] border shadow-xs ${
                            isMatched
                              ? 'bg-white hover:bg-indigo-50/50 border-indigo-200 hover:border-indigo-400 text-slate-800 shadow-indigo-100/40'
                              : 'bg-amber-50 hover:bg-amber-100/50 border-amber-200 hover:border-amber-300 text-amber-800 shadow-amber-100/40'
                          }`}
                        >
                          <div className="w-full flex justify-end">
                            {isMatched ? (
                              <span className="w-4.5 h-4.5 rounded-full bg-emerald-500 text-white flex items-center justify-center text-[9px] font-black shadow-2xs" title="พบวิชานี้ในระบบแล้ว">
                                ✓
                              </span>
                            ) : (
                              <span className="w-4.5 h-4.5 rounded-full bg-amber-500 text-white flex items-center justify-center text-[9px] font-black shadow-2xs" title="วิชาใหม่ (คลิกเพื่อดูคู่มือวิธีเพิ่ม)">
                                ?
                              </span>
                            )}
                          </div>
                          <div className="font-bold text-[12px] leading-tight mt-0.5 line-clamp-2">
                            {slot.subject}
                          </div>
                          <div className="text-[10px] font-semibold opacity-75 mt-1 bg-slate-100 px-2 py-0.5 rounded-md text-slate-600">
                            {slot.class === 'all' ? 'ทุกชั้นเรียน' : `ห้อง ${slot.class}`}
                          </div>
                        </button>
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      
      {/* Timetable Legend */}
      <div className="flex flex-wrap items-center justify-center gap-6 mt-4 pt-4 border-t border-slate-200/60 text-xs font-semibold text-slate-500">
        <span className="flex items-center gap-2">
          <span className="w-3.5 h-3.5 rounded-full bg-emerald-500 flex items-center justify-center text-white text-[8px]">✓</span> วิชาที่มีอยู่ในระบบ Sheets (คลิกเพื่อเริ่มเช็คชื่อทันที)
        </span>
        <span className="flex items-center gap-2">
          <span className="w-3.5 h-3.5 rounded-full bg-amber-500 flex items-center justify-center text-white text-[8px]">?</span> วิชาใหม่/ซ่อมเสริม (คลิกเพื่อดูคำแนะนำการเพิ่มวิชา)
        </span>
      </div>
    </div>
  );
};

export default function ClassroomManager() {
  const envSheetId = import.meta.env.VITE_SHEET_ID;
  const initialSheetUrl = envSheetId ? `https://docs.google.com/spreadsheets/d/${envSheetId}/edit` : (localStorage.getItem('sheetUrl') || '');
  
  const [sheetUrl, setSheetUrl] = useState(initialSheetUrl);
  const [showSetup, setShowSetup] = useState(!initialSheetUrl);
  
  // Auth State
  const [showLogin, setShowLogin] = useState(false);
  const [user, setUser] = useState(null); // null if not logged in. { username, role: 'Teacher' }
  const [loginForm, setLoginForm] = useState({ username: '', password: '' });
  const [loginError, setLoginError] = useState('');

  // Main UI States
  const [activeTab, setActiveTab] = useState('summary'); // Default to summary for public
  const [globalClass, setGlobalClass] = useState('all');
  
  // Data States
  const [attendanceData, setAttendanceData] = useState([]);
  const [assignmentData, setAssignmentData] = useState([]);
  const [studentsData, setStudentsData] = useState([]);
  const [subjectsData, setSubjectsData] = useState([]);
  
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  
  // ---------------- ATTENDANCE STATES ----------------
  const [selectedDate, setSelectedDate] = useState(toLocalDateString(new Date()));
  const [formData, setFormData] = useState({ subject: '' });
  const [weeklyAttendance, setWeeklyAttendance] = useState({});
  const [loadedKey, setLoadedKey] = useState('');
  const [timetableAlert, setTimetableAlert] = useState({ isOpen: false, subject: '', className: '' });
  const [excludedDates, setExcludedDates] = useState([]);

  // ---------------- CROSS-CLASS SUBJECT STATES (ชุมนุม / ลูกเสือ) ----------------
  const [selectedCrossStudents, setSelectedCrossStudents] = useState(new Set()); // Names selected for ชุมนุม
  const [showStudentPicker, setShowStudentPicker] = useState(false); // Show/hide student picker panel

  const toggleDateExclusion = useCallback((dateStr) => {
    setExcludedDates(prev => {
      if (prev.includes(dateStr)) {
        return prev.filter(d => d !== dateStr);
      } else {
        return [...prev, dateStr];
      }
    });
  }, []);





  // ---------------- ASSIGNMENT STATES ----------------
  const [assignForm, setAssignForm] = useState({ subject: '', name: '', dueDate: toLocalDateString(new Date()) });
  const [assignStatusList, setAssignStatusList] = useState({});

  // ---------------- SUMMARY STATES ----------------
  const [summaryClass, setSummaryClass] = useState('');
  const [summaryStudent, setSummaryStudent] = useState('');
  const extractSheetId = (url) => {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
  };

  const fetchSheetData = useCallback(async (sheetId, sheetName, silent = false) => {
    try {
      const apiKey = import.meta.env.VITE_GOOGLE_API_KEY;
      if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
        throw new Error('กรุณาตั้งค่า VITE_GOOGLE_API_KEY ใน .env.local');
      }
      const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${sheetName}?key=${apiKey}`;
      const response = await fetch(apiUrl);
      if (!response.ok) throw new Error(`ไม่พบแผ่นงาน (Sheet) ชื่อ "${sheetName}"`);
      
      const data = await response.json();
      const rows = data.values || [];
      if (rows.length === 0) return [];
      
      const headers = rows[0] || [];
      return rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((header, i) => {
          obj[header.trim()] = row[i] ? row[i].trim() : '';
        });
        return obj;
      });
    } catch (error) {
      console.error('Error fetching data:', error);
      if (!silent) alert('❌ ' + error.message);
      return [];
    }
  }, []);

  const loadAllData = useCallback(async (sheetId) => {
    setLoading(true);
    const [attendance, assignments, students, subjects] = await Promise.all([
      fetchSheetData(sheetId, 'Attendance', true),
      fetchSheetData(sheetId, 'Assignments', true),
      fetchSheetData(sheetId, 'Students', true),
      fetchSheetData(sheetId, 'Subjects', true)
    ]);
    
    setAttendanceData(attendance);
    setAssignmentData(assignments);
    setStudentsData(students);
    setSubjectsData(subjects);
    setLoading(false);
  }, [fetchSheetData]);

  // Fetch data on initial load
  useEffect(() => {
    if (sheetUrl && !showSetup) {
      const sheetId = extractSheetId(sheetUrl);
      if (sheetId) {
        Promise.resolve().then(() => loadAllData(sheetId));
      }
    }
  }, [sheetUrl, showSetup, loadAllData]);

  // Login handler
  const handleLogin = async (e) => {
    e.preventDefault();
    setLoading(true);
    setLoginError('');
    
    const sheetId = extractSheetId(sheetUrl);
    if (!sheetId) {
      setLoginError('URL ของ Google Sheets ไม่ถูกต้อง');
      setLoading(false);
      return;
    }

    try {
      const usersSheet = await fetchSheetData(sheetId, 'Users', true);
      if (usersSheet.length === 0) {
        setLoginError('ไม่พบข้อมูลใน Sheet "Users" หรือยังไม่ได้สร้าง Sheet นี้');
        setLoading(false);
        return;
      }

      // Check credentials
      const foundUser = usersSheet.find(u => 
        (u.Username || u['ชื่อผู้ใช้']) === loginForm.username && 
        (u.Password || u['รหัสผ่าน']) === loginForm.password
      );

      if (foundUser) {
        setUser({ username: loginForm.username, role: 'Teacher' });
        setShowLogin(false);
        setActiveTab('attendance'); // Go to management view on login
      } else {
        setLoginError('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
      }
    } catch {
      setLoginError('เกิดข้อผิดพลาดในการเชื่อมต่อ');
    }
    setLoading(false);
  };

  const handleLogout = () => {
    setUser(null);
    setLoginForm({ username: '', password: '' });
    setActiveTab('summary'); // Return to public view
  };

  const handleSetupSheet = (e) => {
    e.preventDefault();
    if (sheetUrl) {
      localStorage.setItem('sheetUrl', sheetUrl);
      setShowSetup(false);
    }
  };

  // ---------------- DERIVED DATA ----------------
  const normalizedStudents = useMemo(() => {
    return studentsData
      .map(normalizeRow)
      .filter(s => s.name && s.name.length > 1); // กรองเฉพาะคนที่มีชื่อจริงเท่านั้น
  }, [studentsData]);
  
  const classList = useMemo(() => {
    return [...new Set(normalizedStudents.map(s => s.className).filter(c => c !== '-'))].sort();
  }, [normalizedStudents]);

  const subjectList = useMemo(() => {
    return [...new Set(subjectsData.map(normalizeRow).map(s => s.subject).filter(s => s !== '-'))];
  }, [subjectsData]);

  const filteredSubjectList = useMemo(() => {
    if (globalClass === 'all') return subjectList;
    
    // For a specific class, get matching subjects from the sheet
    const classSpecific = subjectList.filter(sub => {
      if (sub.includes(globalClass)) return true;
      const hasOtherClass = classList.some(c => c !== globalClass && sub.includes(c));
      return !hasOtherClass;
    });
    
    // Also include cross-class subjects (ชุมนุม, ลูกเสือ) from the sheet that don't have a specific class
    const crossClassFromSheet = subjectList.filter(sub => isCrossClassSubject(sub));
    
    // Merge without duplicates
    const merged = [...new Set([...classSpecific, ...crossClassFromSheet])];
    return merged;
  }, [subjectList, globalClass, classList]);

  const handleTimetableCellClick = useCallback((dayName, slotInfo) => {
    const dayOffsets = {
      'จันทร์': 0,
      'อังคาร': 1,
      'พุธ': 2,
      'พฤหัสบดี': 3,
      'ศุกร์': 4
    };
    const offset = dayOffsets[dayName];
    if (offset === undefined) return;
    
    const currentRefDate = selectedDate ? new Date(selectedDate) : new Date();
    const currentMonday = getMonday(currentRefDate);
    const targetDate = new Date(currentMonday);
    targetDate.setDate(currentMonday.getDate() + offset);
    const targetDateStr = toLocalDateString(targetDate);
    
    setSelectedDate(targetDateStr);
    
    const targetClass = slotInfo.class;
    if (targetClass !== 'all') {
      setGlobalClass(targetClass);
    }
    
    const resolvedSub = resolveSubjectFromTimetable(slotInfo.subject, targetClass, subjectList);
    
    if (resolvedSub) {
      setFormData({ subject: resolvedSub });
      setActiveTab('attendance');
    } else {
      setFormData({ subject: '' });
      setTimetableAlert({
        isOpen: true,
        subject: slotInfo.subject,
        className: targetClass
      });
    }
    setAssignForm(prev => ({ ...prev, subject: '' }));
  }, [selectedDate, subjectList]);

  const activeYear = useMemo(() => {
    if (!selectedDate) return new Date().getFullYear();
    return new Date(selectedDate).getFullYear();
  }, [selectedDate]);

  const activeMonthIndex = useMemo(() => {
    if (!selectedDate) return new Date().getMonth();
    return new Date(selectedDate).getMonth();
  }, [selectedDate]);

  const teachingDays = useMemo(() => {
    if (!formData.subject) return [];
    
    const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    const matchedDays = [];
    const isCross = isCrossClassSubject(formData.subject);
    
    if (!isCross && globalClass === 'all') return [];
    
    for (const [dayName, daySchedule] of Object.entries(SCHEDULE)) {
      for (const slot of Object.values(daySchedule)) {
        // For cross-class subjects: match slots where class='all'
        // For normal subjects: match slots where class=globalClass
        const classMatch = isCross ? (slot.class === 'all') : (slot.class === globalClass);
        if (!classMatch) continue;
        
        const resolved = resolveSubjectFromTimetable(slot.subject, slot.class, subjectList);
        if (resolved === formData.subject) {
          const dayIdx = dayNames.indexOf(dayName);
          if (dayIdx !== -1 && !matchedDays.includes(dayIdx)) {
            matchedDays.push(dayIdx);
          }
        }
        
        // For cross-class: also match by keyword if no exact resolve found
        if (isCross && !resolved) {
          const subjectKeyword = CROSS_CLASS_KEYWORDS.find(kw => formData.subject.includes(kw));
          if (subjectKeyword && slot.subject.includes(subjectKeyword)) {
            const dayIdx = dayNames.indexOf(dayName);
            if (dayIdx !== -1 && !matchedDays.includes(dayIdx)) {
              matchedDays.push(dayIdx);
            }
          }
        }
      }
    }
    return matchedDays.sort();
  }, [globalClass, formData.subject, subjectList]);

  const getDatesForWeekdayInMonth = useCallback((year, monthIndex, weekdayIndex) => {
    const dates = [];
    const date = new Date(year, monthIndex, 1);
    while (date.getDay() !== weekdayIndex) {
      date.setDate(date.getDate() + 1);
    }
    while (date.getMonth() === monthIndex) {
      dates.push(toLocalDateString(date));
      date.setDate(date.getDate() + 7);
    }
    return dates;
  }, []);

  const activeGridDates = useMemo(() => {
    if (!formData.subject) return [];
    const isCross = isCrossClassSubject(formData.subject);
    if (!isCross && globalClass === 'all') return [];
    
    const targetDays = teachingDays.length > 0 ? teachingDays : [1, 2, 3, 4, 5];
    
    let dates = [];
    targetDays.forEach(dayIdx => {
      const dayDates = getDatesForWeekdayInMonth(activeYear, activeMonthIndex, dayIdx);
      dates = dates.concat(dayDates);
    });
    
    dates.sort((a, b) => new Date(a) - new Date(b));
    
    // กรองวันที่ในอนาคตออกเพื่อไม่ให้มีคอลัมน์ล่วงหน้าและป้องกันการบันทึกข้อมูลล่วงหน้า
    const todayStr = toLocalDateString(new Date());
    dates = dates.filter(d => d <= todayStr);
    
    return dates;
  }, [teachingDays, activeYear, activeMonthIndex, globalClass, formData.subject, getDatesForWeekdayInMonth]);

  const currentKey = `${globalClass}|${formData.subject}|${activeYear}-${activeMonthIndex}`;

  // Active Class Students for Checklists
  const activeClassStudents = useMemo(() => {
    const subject = formData.subject;
    
    // ลูกเสือ: auto-load ป.4-ม.3 all students
    if (isScoutSubject(subject)) {
      return normalizedStudents.filter(s => SCOUT_CLASSES.includes(s.className));
    }
    
    // ชุมนุม: use teacher-selected students
    if (isClubSubject(subject)) {
      if (selectedCrossStudents.size === 0) return [];
      return normalizedStudents.filter(s => selectedCrossStudents.has(s.name));
    }
    
    // Normal class subject
    if (globalClass === 'all') return [];
    return normalizedStudents.filter(s => s.className === globalClass);
  }, [normalizedStudents, globalClass, formData.subject, selectedCrossStudents]);

  // Students list for Summary Dropdown
  const summaryStudentList = useMemo(() => {
    if (!summaryClass) return [];
    return normalizedStudents.filter(s => s.className === summaryClass).map(s => s.name).sort();
  }, [normalizedStudents, summaryClass]);

  const createInitialStatusLists = (students) => {
    const initialAtt = {};
    const initialAssign = {};
    students.forEach(s => {
      initialAtt[s.name] = 'present';
      initialAssign[s.name] = 'pending';
    });

    return { initialAtt, initialAssign };
  };

  const handleGlobalClassChange = (className) => {
    setGlobalClass(className);
    const students = className === 'all' ? [] : normalizedStudents.filter(s => s.className === className);
    const { initialAssign } = createInitialStatusLists(students);
    setAssignStatusList(initialAssign);
    setFormData(prev => ({ ...prev, subject: '' }));
    setAssignForm(prev => ({ ...prev, subject: '' }));
    setLoadedKey(''); // Force reload weekly attendance
    // Reset cross-class selections
    setSelectedCrossStudents(new Set());
    setShowStudentPicker(false);
  };

  const getAttendanceStatus = (name, date) => weeklyAttendance[name]?.[date] || 'present';
  const getAssignmentStatus = (name) => assignStatusList[name] || 'pending';

  // Normalized Histories
  const normalizedAttendanceHistory = useMemo(() => {
    return attendanceData.map(row => {
      const data = normalizeRow(row);
      const studentMatch = normalizedStudents.find(s => s.name === data.name);
      return { ...data, className: studentMatch && studentMatch.className !== '-' ? studentMatch.className : 'ไม่ระบุ' };
    });
  }, [attendanceData, normalizedStudents]);

  const filteredAttendanceHistory = useMemo(() => {
    if (globalClass === 'all') return normalizedAttendanceHistory;
    return normalizedAttendanceHistory.filter(r => r.className === globalClass);
  }, [normalizedAttendanceHistory, globalClass]);

  const normalizedAssignmentHistory = useMemo(() => {
    return assignmentData.map(row => {
      const data = normalizeRow(row);
      const studentMatch = normalizedStudents.find(s => s.name === data.name);
      return { ...data, className: studentMatch && studentMatch.className !== '-' ? studentMatch.className : 'ไม่ระบุ' };
    });
  }, [assignmentData, normalizedStudents]);

  const filteredAssignmentHistory = useMemo(() => {
    if (globalClass === 'all') return normalizedAssignmentHistory;
    return normalizedAssignmentHistory.filter(r => r.className === globalClass);
  }, [normalizedAssignmentHistory, globalClass]);

  // Options for past attendance (grouped by month)
  const pastAttendanceOptions = useMemo(() => {
    const options = [];
    const seen = new Set();
    filteredAttendanceHistory.forEach(record => {
      if (!record.date || record.date === '-' || record.subject === '-') return;
      const dateObj = new Date(record.date);
      if (isNaN(dateObj.getTime())) return;
      const monthStr = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-01`;
      const key = `${monthStr}|${record.subject}`;
      if (!seen.has(key)) {
        seen.add(key);
        options.push({ monthStr, subject: record.subject });
      }
    });
    return options.sort((a, b) => new Date(b.monthStr) - new Date(a.monthStr));
  }, [filteredAttendanceHistory]);

  // Load monthly attendance when currentKey or history changes
  useEffect(() => {
    let timerId;
    const isCross = isCrossClassSubject(formData.subject);
    const shouldReset = (!isCross && globalClass === 'all') || !formData.subject || !selectedDate;
    
    if (shouldReset) {
      timerId = setTimeout(() => {
        setWeeklyAttendance({});
        setExcludedDates([]);
        setLoadedKey('');
      }, 0);
    } else if (loadedKey !== currentKey) {
      // For cross-class subjects (ชุมนุม/ลูกเสือ), use full attendance history
      const attHistory = isCross ? normalizedAttendanceHistory : filteredAttendanceHistory;
      
      const newWeeklyAtt = {};
      activeClassStudents.forEach(student => {
        newWeeklyAtt[student.name] = {};
        activeGridDates.forEach(dateStr => {
          const record = attHistory.find(r => 
            r.name === student.name && 
            r.subject === formData.subject && 
            r.date === dateStr
          );
          newWeeklyAtt[student.name][dateStr] = record ? record.status : 'present';
        });
      });

      // Check if there are any saved records for this subject in the active grid dates
      const hasAnySavedRecords = attHistory.some(r => 
        r.subject === formData.subject && 
        activeGridDates.includes(r.date)
      );

      let initialExcluded = [];
      if (hasAnySavedRecords) {
        initialExcluded = activeGridDates.filter(dateStr => 
          !attHistory.some(r => 
            r.subject === formData.subject && 
            r.date === dateStr
          )
        );
      }

      timerId = setTimeout(() => {
        setWeeklyAttendance(newWeeklyAtt);
        setExcludedDates(initialExcluded);
        setLoadedKey(currentKey);
      }, 0);
    }
    return () => {
      if (timerId) clearTimeout(timerId);
    };
  }, [currentKey, activeClassStudents, activeGridDates, filteredAttendanceHistory, normalizedAttendanceHistory, loadedKey, globalClass, formData.subject, selectedDate]);


  // Options for past assignments
  const pastAssignmentOptions = useMemo(() => {
    const options = [];
    const seen = new Set();
    filteredAssignmentHistory.forEach(record => {
      const key = `${record.assignment}|${record.subject}|${record.dueDate}`;
      if (!seen.has(key) && record.assignment !== '-' && record.subject !== '-') {
        seen.add(key);
        options.push({ assignment: record.assignment, subject: record.subject, dueDate: record.dueDate });
      }
    });
    return options.sort((a, b) => new Date(b.dueDate) - new Date(a.dueDate));
  }, [filteredAssignmentHistory]);


  // ---------------- HANDLERS ----------------
  const handleStatusChange = (name, date, status) => {
    setWeeklyAttendance(prev => ({
      ...prev,
      [name]: {
        ...(prev[name] || {}),
        [date]: status
      }
    }));
  };

  const toggleAttendanceStatus = (name, date) => {
    const currentStatus = getAttendanceStatus(name, date);
    let nextStatus;
    if (currentStatus === 'present') {
      nextStatus = 'absent';
    } else if (currentStatus === 'absent') {
      nextStatus = 'late';
    } else {
      nextStatus = 'present';
    }
    handleStatusChange(name, date, nextStatus);
  };

  const handleMarkAllDayPresent = (dateStr) => {
    setWeeklyAttendance(prev => {
      const updated = { ...prev };
      activeClassStudents.forEach(student => {
        if (!updated[student.name]) {
          updated[student.name] = {};
        }
        updated[student.name][dateStr] = 'present';
      });
      return updated;
    });
  };

  const handleAssignStatusChange = (name, status) => setAssignStatusList(prev => ({ ...prev, [name]: status }));

  const handleLoadPastAttendance = (val) => {
    if (!val) {
      setSelectedDate(toLocalDateString(new Date()));
      setFormData({ subject: '' });
      setLoadedKey('');
      return;
    }
    const [mStr, subject] = val.split('|');
    setSelectedDate(mStr);
    setFormData({ subject });
    setLoadedKey('');
  };

  const handleLoadPastAssignment = (val) => {
    if (!val) {
      setAssignForm({ subject: '', name: '', dueDate: toLocalDateString(new Date()) });
      const initialAssign = {};
      activeClassStudents.forEach(s => initialAssign[s.name] = 'pending');
      setAssignStatusList(initialAssign);
      return;
    }
    const [assignment, subject, dueDate] = val.split('|');
    setAssignForm({ name: assignment, subject, dueDate });
    
    const pastRecords = filteredAssignmentHistory.filter(r => r.assignment === assignment && r.subject === subject && r.dueDate === dueDate);
    const newAssignList = {};
    activeClassStudents.forEach(s => {
      const record = pastRecords.find(r => r.name === s.name);
      newAssignList[s.name] = record ? record.status : 'pending';
    });
    setAssignStatusList(newAssignList);
  };

  const handleSaveBulkAttendance = async (e) => {
    e.preventDefault();
    const isCross = isCrossClassSubject(formData.subject);
    if (!isCross && globalClass === 'all') return alert('กรุณาเลือกชั้นเรียนก่อน');
    if (!formData.subject) return alert('กรุณาเลือกวิชาก่อนบันทึก');
    if (isClubSubject(formData.subject) && selectedCrossStudents.size === 0) {
      return alert('กรุณาเลือกรายชื่อนักเรียนชุมนุมก่อนบันทึก (คลิกปุ่ม "เลือกรายชื่อนักเรียน")');
    }
    
    const scriptUrl = import.meta.env.VITE_GOOGLE_APP_SCRIPT_URL;
    
    // Validate URL
    if (!scriptUrl || !scriptUrl.startsWith('https://script.google.com')) {
      console.error('Invalid or missing VITE_GOOGLE_APP_SCRIPT_URL:', scriptUrl);
      return alert('⚠️ URL สำหรับบันทึกข้อมูลไม่ถูกต้อง หรือยังไม่ได้ตั้งค่าใน Vercel\n\nโปรดตรวจสอบว่าได้ตั้งค่า VITE_GOOGLE_APP_SCRIPT_URL ให้เป็น URL ที่ขึ้นต้นด้วย https://script.google.com ...');
    }

    const sheetId = extractSheetId(sheetUrl);
    if (!sheetId) {
      return alert('❌ URL ของ Google Sheets ไม่ถูกต้อง ไม่สามารถดึง Sheet ID ได้');
    }

    setIsSaving(true);
    try {
      const validStudents = activeClassStudents.filter(s => s.name && String(s.name).trim().length > 1);
      
      const datesToSave = activeGridDates.filter(d => !excludedDates.includes(d));
      const payloadData = [];
      validStudents.forEach(student => {
        datesToSave.forEach(dateStr => {
          const status = getAttendanceStatus(student.name, dateStr);
          payloadData.push(createAttendancePayloadRow({
            student,
            subject: formData.subject,
            date: dateStr,
            status,
            className: globalClass
          }));
        });
      });

      if (payloadData.length === 0) {
        setIsSaving(false);
        return alert('❌ ไม่พบรายชื่อนักเรียนหรือวันที่ที่จะบันทึก (โปรดมั่นใจว่าไม่ได้เลือกข้ามทุกวัน)');
      }

      const confirmMsg = `ยืนยันบันทึกข้อมูลการเช็คชื่อเข้าเรียน\nห้อง: ${globalClass}\nวิชา: ${formData.subject}\nประจำเดือน: ${thaiMonthName(activeMonthIndex)} ${activeYear + 543} (จำนวน ${datesToSave.length} วันสอนจากทั้งหมด ${activeGridDates.length} วัน)\nจำนวนนักเรียน: ${validStudents.length} คน (รวมทั้งหมด ${payloadData.length} แถว)\n\nต้องการดำเนินการต่อหรือไม่?`;
      
      if (!window.confirm(confirmMsg)) {
        setIsSaving(false);
        return;
      }

      console.log('--- Bulk Payload Ready ---', payloadData);

      await fetch(scriptUrl, {
        method: 'POST',
        mode: 'no-cors', // ข้ามปัญหา CORS โดยการไม่รออ่าน response
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ 
          action: 'attendance', 
          sheetId: sheetId, 
          data: payloadData,
          mode: 'overwrite',
          className: globalClass
        })
      });
      
      alert('✅ ส่งข้อมูลการเข้าเรียนรายสัปดาห์เรียบร้อยแล้ว!\n(โปรดตรวจสอบความถูกต้องใน Google Sheets อีกครั้ง)');
      setLoadedKey('');
      setTimeout(() => loadAllData(sheetId), 1500);
    } catch (error) {
      console.error('Save Error:', error);
      alert('❌ บันทึกไม่สำเร็จ: ' + error.message);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSaveBulkAssignment = async (e) => {
    e.preventDefault();
    if (globalClass === 'all') return alert('กรุณาเลือกชั้นเรียนก่อน');
    if (!assignForm.subject || !assignForm.name) return alert('กรุณาเลือกวิชาและตั้งชื่องานก่อนบันทึก');

    const scriptUrl = import.meta.env.VITE_GOOGLE_APP_SCRIPT_URL;

    // Validate URL
    if (!scriptUrl || !scriptUrl.startsWith('https://script.google.com')) {
      console.error('Invalid or missing VITE_GOOGLE_APP_SCRIPT_URL:', scriptUrl);
      return alert('⚠️ URL สำหรับบันทึกข้อมูลไม่ถูกต้อง หรือยังไม่ได้ตั้งค่าใน Vercel\n\nโปรดตรวจสอบว่าได้ตั้งค่า VITE_GOOGLE_APP_SCRIPT_URL ให้เป็น URL ที่ขึ้นต้นด้วย https://script.google.com ...');
    }

    const sheetId = extractSheetId(sheetUrl);
    if (!sheetId) {
      return alert('❌ URL ของ Google Sheets ไม่ถูกต้อง ไม่สามารถดึง Sheet ID ได้');
    }

    setIsSaving(true);
    try {
      const validStudents = activeClassStudents.filter(s => s.name && String(s.name).trim().length > 1);

      const payloadData = validStudents.map(student => createAssignmentPayloadRow({
        student,
        subject: assignForm.subject,
        assignment: assignForm.name,
        dueDate: assignForm.dueDate,
        status: getAssignmentStatus(student.name),
        className: globalClass
      }));

      if (payloadData.length === 0) {
        setIsSaving(false);
        return alert('❌ ไม่พบรายชื่อนักเรียนที่จะบันทึก');
      }

      // แสดงตัวอย่างข้อมูลเพื่อ Debug
      const sample = payloadData[0];
      const confirmMsg = `ยืนยันบันทึกส่งงาน ${payloadData.length} คน\nงาน: ${sample.Assignment}\nวิชา: ${sample.Subject}\nตัวอย่างชื่อคนแรก: ${sample.Name}\n\nต้องการดำเนินการต่อหรือไม่?`;

      if (!window.confirm(confirmMsg)) {
        setIsSaving(false);
        return;
      }

      console.log('Final Assignment Payload:', payloadData);
      console.log('Attempting to save assignment to (no-cors mode):', scriptUrl.substring(0, 45) + '...');

      await fetch(scriptUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ 
          action: 'assignment', 
          sheetId: sheetId, 
          data: payloadData,
          mode: 'overwrite',
          className: globalClass
        })
      });
      
      alert('✅ ส่งข้อมูลการส่งงานเรียบร้อยแล้ว!\n(โปรดตรวจสอบความถูกต้องใน Google Sheets อีกครั้ง)');
      setTimeout(() => loadAllData(sheetId), 1500);
    } catch (error) {
      console.error('Save Error:', error);
      alert('❌ บันทึกไม่สำเร็จ: ' + error.message);
    } finally {
      setIsSaving(false);
    }
  };

  // ---------------- STATS CALC ----------------
  const getAttendanceStats = () => {
    const present = filteredAttendanceHistory.filter(d => d.status === 'present').length;
    const absent = filteredAttendanceHistory.filter(d => d.status === 'absent').length;
    const late = filteredAttendanceHistory.filter(d => d.status === 'late').length;
    return [
      { name: 'มาเรียนรวม', value: present, fill: '#10b981', bg: 'bg-emerald-500/10', text: 'text-emerald-500', icon: <CheckCircle className="text-emerald-500" size={24} /> },
      { name: 'ขาดเรียนรวม', value: absent, fill: '#ef4444', bg: 'bg-rose-500/10', text: 'text-rose-500', icon: <AlertCircle className="text-rose-500" size={24} /> },
      { name: 'มาสายรวม', value: late, fill: '#f59e0b', bg: 'bg-amber-500/10', text: 'text-amber-500', icon: <Clock className="text-amber-500" size={24} /> }
    ];
  };

  const getAssignmentStats = () => {
    const completed = filteredAssignmentHistory.filter(d => d.status === 'completed').length;
    const pending = filteredAssignmentHistory.filter(d => d.status === 'pending').length;
    const overdue = filteredAssignmentHistory.filter(d => d.status === 'overdue').length;
    return [
      { name: 'เสร็จสิ้น', value: completed, fill: '#10b981', bg: 'bg-emerald-500/10', text: 'text-emerald-500', icon: <CheckCircle className="text-emerald-500" size={24} /> },
      { name: 'กำลังทำ', value: pending, fill: '#3b82f6', bg: 'bg-blue-500/10', text: 'text-blue-500', icon: <Clock className="text-blue-500" size={24} /> },
      { name: 'เลยกำหนด', value: overdue, fill: '#ef4444', bg: 'bg-rose-500/10', text: 'text-rose-500', icon: <AlertCircle className="text-rose-500" size={24} /> }
    ];
  };

  const summaryAttStats = useMemo(() => {
    const data = summaryClass ? normalizedAttendanceHistory.filter(h => h.className === summaryClass) : normalizedAttendanceHistory;
    const present = data.filter(d => d.status === 'present').length;
    const absent = data.filter(d => d.status === 'absent').length;
    const late = data.filter(d => d.status === 'late').length;
    return [
      { name: 'มาเรียน', value: present, fill: '#10b981' },
      { name: 'ขาดเรียน', value: absent, fill: '#ef4444' },
      { name: 'มาสาย', value: late, fill: '#f59e0b' }
    ].filter(s => s.value > 0);
  }, [normalizedAttendanceHistory, summaryClass]);

  const summaryAssignStats = useMemo(() => {
    const data = summaryClass ? normalizedAssignmentHistory.filter(h => h.className === summaryClass) : normalizedAssignmentHistory;
    const completed = data.filter(d => d.status === 'completed').length;
    const pending = data.filter(d => d.status === 'pending').length;
    const overdue = data.filter(d => d.status === 'overdue').length;
    return [
      { name: 'ส่งแล้ว', value: completed, fill: '#10b981' },
      { name: 'กำลังทำ', value: pending, fill: '#3b82f6' },
      { name: 'เลยกำหนด', value: overdue, fill: '#ef4444' }
    ].filter(s => s.value > 0);
  }, [normalizedAssignmentHistory, summaryClass]);

  const exportAttendanceExcel = () => {
    const targetStudents = globalClass === 'all' ? normalizedStudents : activeClassStudents;
    const historyData = globalClass === 'all' ? normalizedAttendanceHistory : filteredAttendanceHistory;

    if (historyData.length === 0) return alert('ไม่มีข้อมูลสำหรับส่งออก');
    
    const excelData = targetStudents.map(student => {
      const studentHistory = historyData.filter(h => h.name === student.name);
      const present = studentHistory.filter(h => h.status === 'present').length;
      const absent = studentHistory.filter(h => h.status === 'absent').length;
      const late = studentHistory.filter(h => h.status === 'late').length;
      
      return {
        'ชั้นเรียน': student.className,
        'เลขที่': student.studentId,
        'ชื่อ-สกุล': student.name,
        'มาเรียน (ครั้ง)': present,
        'ขาดเรียน (ครั้ง)': absent,
        'มาสาย (ครั้ง)': late,
        'รวมเช็คชื่อทั้งหมด (ครั้ง)': present + absent + late
      };
    });
    
    excelData.sort((a, b) => {
      if (a['ชั้นเรียน'] !== b['ชั้นเรียน']) return String(a['ชั้นเรียน']).localeCompare(String(b['ชั้นเรียน']));
      return (parseInt(a['เลขที่']) || 0) - (parseInt(b['เลขที่']) || 0);
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    ws['!cols'] = [{ wch: 10 }, { wch: 10 }, { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance Summary");
    XLSX.writeFile(wb, `สรุปการเข้าเรียน_${globalClass === 'all' ? 'รวม' : `ห้อง_${globalClass}`}.xlsx`);
  };

  const exportAssignmentExcel = () => {
    const targetStudents = globalClass === 'all' ? normalizedStudents : activeClassStudents;
    const historyData = globalClass === 'all' ? normalizedAssignmentHistory : filteredAssignmentHistory;

    if (historyData.length === 0) return alert('ไม่มีข้อมูลสำหรับส่งออก');
    
    const excelData = targetStudents.map(student => {
      const studentHistory = historyData.filter(h => h.name === student.name);
      const completed = studentHistory.filter(h => h.status === 'completed').length;
      const pending = studentHistory.filter(h => h.status === 'pending').length;
      const overdue = studentHistory.filter(h => h.status === 'overdue').length;
      
      return {
        'ชั้นเรียน': student.className,
        'เลขที่': student.studentId,
        'ชื่อ-สกุล': student.name,
        'ส่งแล้ว (ชิ้น)': completed,
        'กำลังทำ (ชิ้น)': pending,
        'เลยกำหนด (ชิ้น)': overdue,
        'รวมงานทั้งหมด (ชิ้น)': completed + pending + overdue
      };
    });
    
    excelData.sort((a, b) => {
      if (a['ชั้นเรียน'] !== b['ชั้นเรียน']) return String(a['ชั้นเรียน']).localeCompare(String(b['ชั้นเรียน']));
      return (parseInt(a['เลขที่']) || 0) - (parseInt(b['เลขที่']) || 0);
    });

    const ws = XLSX.utils.json_to_sheet(excelData);
    ws['!cols'] = [{ wch: 10 }, { wch: 10 }, { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Assignment Summary");
    XLSX.writeFile(wb, `สรุปการส่งงาน_${globalClass === 'all' ? 'รวม' : `ห้อง_${globalClass}`}.xlsx`);
  };

  // ---------------- INDIVIDUAL SUMMARY CALC ----------------
  const studentAttHistory = useMemo(() => {
    if (!summaryStudent) return [];
    return normalizedAttendanceHistory.filter(h => h.name === summaryStudent).sort((a,b) => new Date(b.date) - new Date(a.date));
  }, [summaryStudent, normalizedAttendanceHistory]);

  const studentAssignHistory = useMemo(() => {
    if (!summaryStudent) return [];
    return normalizedAssignmentHistory.filter(h => h.name === summaryStudent).sort((a,b) => new Date(b.dueDate) - new Date(a.dueDate));
  }, [summaryStudent, normalizedAssignmentHistory]);

  // ================= RENDER SETUP SCREEN =================
  if (showSetup) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4 relative">
        <DecorativeBackground />
        <div className="glass-card rounded-[2.5rem] p-10 max-w-xl w-full animate-fade-in-up border-white/60 relative overflow-hidden">
          <div className="absolute top-0 right-0 p-8 opacity-10"><Sparkles size={120} /></div>
          <div className="relative z-10">
            <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-gradient-to-tr from-indigo-500 to-purple-500 text-white shadow-lg shadow-indigo-200 mb-8"><LayoutDashboard size={32} /></div>
            <h1 className="text-4xl font-extrabold mb-3 text-slate-800 tracking-tight">Classroom <span className="text-transparent bg-clip-text bg-gradient-to-r from-indigo-600 to-purple-600">Manager</span></h1>
            <p className="text-slate-500 mb-10 text-lg">เชื่อมต่อระบบเข้ากับ Google Sheets ของคุณ</p>
            <form onSubmit={handleSetupSheet} className="space-y-6">
              <div className="space-y-2">
                <label className="block text-sm font-semibold text-slate-700 ml-1">Google Sheet URL</label>
                <div className="relative">
                  <input type="url" value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/..." className="w-full pl-4 pr-12 py-4 bg-white/50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all duration-300 text-slate-700 shadow-sm" required/>
                  <div className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400"><Settings size={20} /></div>
                </div>
              </div>
              <button type="submit" className="w-full bg-slate-900 text-white py-4 rounded-2xl hover:bg-slate-800 font-semibold transition-all duration-300 flex items-center justify-center gap-2 shadow-xl shadow-slate-900/20 hover:-translate-y-1">
                ดำเนินการต่อ <ArrowRight size={18} />
              </button>
            </form>
          </div>
        </div>
      </div>
    );
  }

  // ================= RENDER LOGIN SCREEN =================
  if (showLogin) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4 relative">
        <DecorativeBackground />
        
        <button onClick={() => setShowLogin(false)} className="absolute top-6 left-6 md:top-10 md:left-10 text-slate-500 hover:text-indigo-600 flex items-center gap-2 font-bold transition-colors bg-white/50 px-5 py-2.5 rounded-full shadow-sm">
          <ArrowRight className="rotate-180" size={18} /> กลับหน้าหลัก
        </button>

        <div className="glass-card rounded-[2.5rem] p-10 max-w-md w-full animate-fade-in-up border-white/60 shadow-2xl relative overflow-hidden">
          <div className="absolute -top-10 -right-10 w-40 h-40 bg-indigo-500/10 rounded-full blur-2xl"></div>
          <div className="relative z-10">
            <div className="flex flex-col items-center mb-8">
              <div className="w-20 h-20 rounded-3xl bg-gradient-to-tr from-indigo-500 to-purple-500 text-white shadow-xl shadow-indigo-200/50 flex items-center justify-center mb-6">
                <Lock size={36} />
              </div>
              <h1 className="text-2xl font-extrabold text-slate-800">สำหรับคุณครู</h1>
              <p className="text-slate-500 text-sm mt-2 text-center">เข้าสู่ระบบเพื่อบันทึกและจัดการข้อมูลชั้นเรียน</p>
            </div>

            {loginError && (
              <div className="mb-6 p-4 bg-rose-50 border border-rose-100 rounded-xl text-rose-600 text-sm font-medium flex items-center gap-2 animate-shake">
                <AlertCircle size={18} /> {loginError}
              </div>
            )}

            <form onSubmit={handleLogin} className="space-y-5">
              <div className="space-y-1.5">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider ml-1">ชื่อผู้ใช้ (Username)</label>
                <div className="relative">
                  <input type="text" value={loginForm.username} onChange={e => setLoginForm({...loginForm, username: e.target.value})} className="w-full pl-11 pr-4 py-3.5 bg-white/60 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all text-sm font-medium text-slate-700" placeholder="กรอกชื่อผู้ใช้..." required/>
                  <User size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" />
                </div>
              </div>
              <div className="space-y-1.5">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider ml-1">รหัสผ่าน (Password)</label>
                <div className="relative">
                  <input type="password" value={loginForm.password} onChange={e => setLoginForm({...loginForm, password: e.target.value})} className="w-full pl-11 pr-4 py-3.5 bg-white/60 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all text-sm font-medium text-slate-700" placeholder="••••••••" required/>
                  <Lock size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" />
                </div>
              </div>
              <button type="submit" disabled={loading} className="w-full mt-2 bg-gradient-to-r from-indigo-600 to-purple-600 text-white py-4 rounded-2xl hover:from-indigo-700 hover:to-purple-700 font-bold transition-all shadow-xl shadow-indigo-600/20 hover:-translate-y-1 disabled:opacity-70 disabled:hover:translate-y-0 flex items-center justify-center gap-2">
                {loading ? <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin"></div> : 'เข้าสู่ระบบ'}
              </button>
            </form>
          </div>
        </div>
      </div>
    );
  }

  // ================= RENDER MAIN APP =================
  const isTeacher = user !== null;

  return (
    <div className="min-h-screen pb-20 relative">
      <DecorativeBackground />

      {/* HEADER */}
      <header className="glass-nav sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-500 to-purple-500 flex items-center justify-center text-white shadow-md">
              <LayoutDashboard size={20} />
            </div>
            <h1 className="text-xl font-bold text-slate-800 tracking-tight hidden sm:block">
              Classroom <span className="text-indigo-600">Portal</span>
            </h1>
          </div>
          <div className="flex items-center gap-3 sm:gap-6">
            {isTeacher ? (
              <>
                <div className="hidden sm:flex flex-col items-end">
                  <span className="text-sm font-bold text-slate-700">{user.username}</span>
                  <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-600">คุณครู</span>
                </div>
                <button onClick={handleLogout} className="flex items-center gap-2 px-4 py-2 bg-rose-50 text-rose-600 rounded-full hover:bg-rose-100 transition-all font-medium text-sm">
                  <LogOut size={16} /> ออกจากระบบ
                </button>
              </>
            ) : (
              <button onClick={() => setShowLogin(true)} className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 text-white rounded-full hover:bg-indigo-700 transition-all shadow-md shadow-indigo-200 font-medium text-sm">
                <Lock size={16} /> สำหรับคุณครู
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-8">
        
        {/* TOP BAR: Tabs + Global Class Filter */}
        <div className="flex flex-col sm:flex-row justify-between items-center mb-8 gap-4">
          <div className="inline-flex bg-white/50 backdrop-blur-md p-1.5 rounded-full border border-slate-200/60 shadow-sm w-full sm:w-auto overflow-x-auto">
            <button onClick={() => setActiveTab('summary')} className={`whitespace-nowrap px-6 py-2.5 rounded-full text-sm font-semibold transition-all duration-300 flex items-center gap-2 ${activeTab === 'summary' ? 'bg-white text-emerald-600 shadow-md scale-100' : 'text-slate-500 hover:text-slate-700'}`}>
              <Award size={16} /> ดูสรุปรายบุคคล
            </button>
            {isTeacher && (
              <>
                <button onClick={() => setActiveTab('attendance')} className={`whitespace-nowrap px-6 py-2.5 rounded-full text-sm font-semibold transition-all duration-300 flex items-center gap-2 ${activeTab === 'attendance' ? 'bg-white text-indigo-600 shadow-md scale-100' : 'text-slate-500 hover:text-slate-700'}`}>
                  <Users size={16} /> การเข้าเรียน
                </button>
                <button onClick={() => setActiveTab('assignments')} className={`whitespace-nowrap px-6 py-2.5 rounded-full text-sm font-semibold transition-all duration-300 flex items-center gap-2 ${activeTab === 'assignments' ? 'bg-white text-purple-600 shadow-md scale-100' : 'text-slate-500 hover:text-slate-700'}`}>
                  <BookOpen size={16} /> การส่งงาน
                </button>
              </>
            )}
          </div>

          {isTeacher && activeTab !== 'summary' && (
            <div className="glass-card px-5 py-2.5 rounded-full flex items-center gap-3 border-indigo-100 shadow-sm w-full sm:w-auto">
              <Users size={16} className="text-indigo-500" />
              <span className="text-sm font-semibold text-slate-600 whitespace-nowrap">ชั้นเรียน:</span>
              <select value={globalClass} onChange={(e) => handleGlobalClassChange(e.target.value)} className="w-full sm:w-auto bg-transparent border-none text-indigo-700 font-bold text-sm focus:ring-0 outline-none cursor-pointer pr-4">
                <option value="all">ดูภาพรวมทุกชั้น</option>
                {classList.map(cls => <option key={cls} value={cls}>ห้อง {cls}</option>)}
              </select>
            </div>
          )}
        </div>

        {(loading || isSaving) ? (
          <div className="flex flex-col items-center justify-center py-32 animate-fade-in-up">
            <div className="w-12 h-12 border-4 border-indigo-200 border-t-indigo-600 rounded-full animate-spin mb-4"></div>
            <p className="text-slate-500 font-medium animate-pulse">{isSaving ? 'กำลังบันทึกข้อมูล...' : 'กำลังโหลดข้อมูล...'}</p>
          </div>
        ) : (
          <div className="opacity-100 translate-y-0 transition-all duration-500 ease-out">
            
            {/* ---------------- ATTENDANCE TAB (TEACHER ONLY) ---------------- */}
            {activeTab === 'attendance' && isTeacher && (
              <div className="space-y-8">
                {/* Stats */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {getAttendanceStats().map((stat) => (
                    <div key={stat.name} className="glass-card rounded-[2rem] p-6 hover:-translate-y-1 transition-all duration-300">
                      <div className="flex items-center gap-4">
                        <div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${stat.bg}`}>{stat.icon}</div>
                        <div>
                          <p className="text-slate-500 text-sm font-medium mb-1">{stat.name}</p>
                          <h3 className={`text-3xl font-bold ${stat.text}`}>{stat.value} <span className="text-sm font-normal text-slate-400">ครั้ง</span></h3>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>

                {globalClass === 'all' && !formData.subject ? (
                  <div className="space-y-8">
                    <div className="glass-card rounded-[2rem] p-4 text-center flex flex-col items-center justify-center border-dashed border-2 border-indigo-200/50 bg-indigo-50/20 animate-fade-in-up">
                      <Users size={32} className="text-indigo-500 mb-2" />
                      <h3 className="text-md font-bold text-indigo-900">กรุณาเลือกชั้นเรียนที่แถบด้านบน หรือคลิกจากตารางเรียนครูวีรวัฒน์ด้านล่าง</h3>
                      <p className="text-xs text-indigo-700/70 mt-0.5">ระบบจะเลือกชั้นเรียนและรายวิชาให้โดยอัตโนมัติ</p>
                    </div>
                    <TimetableGrid 
                      subjectList={subjectList}
                      onCellClick={handleTimetableCellClick}
                    />
                  </div>
                ) : (
                  <>
                    <div className="glass-card rounded-[2rem] p-6 flex flex-col gap-4 shadow-sm mb-8 animate-fade-in-up">
                      {pastAttendanceOptions.length > 0 && (
                        <div className="w-full pb-4 border-b border-slate-100">
                          <label className="block text-xs font-semibold text-indigo-500 mb-2 uppercase tracking-wider flex items-center gap-1"><Edit3 size={14}/> เลือกเดือนเดิมเพื่อแก้ไข</label>
                          <select 
                            value={(() => {
                              if (!formData.subject) return "";
                              const currentMonthStr = `${activeYear}-${String(activeMonthIndex + 1).padStart(2, '0')}-01`;
                              const expectedSelectValue = `${currentMonthStr}|${formData.subject}`;
                              const isPastAttendanceExists = pastAttendanceOptions.some(opt => `${opt.monthStr}|${opt.subject}` === expectedSelectValue);
                              return isPastAttendanceExists ? expectedSelectValue : "";
                            })()}
                            onChange={(e) => handleLoadPastAttendance(e.target.value)} 
                            className="w-full px-4 py-3 bg-indigo-50/50 border border-indigo-100 rounded-xl focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm font-medium text-indigo-700 cursor-pointer transition-colors hover:bg-indigo-50"
                          >
                            <option value="">-- สร้างการเช็คชื่อเดือนใหม่ --</option>
                            {pastAttendanceOptions.map((opt, i) => (
                              <option key={i} value={`${opt.monthStr}|${opt.subject}`}>
                                ประจำเดือน {thaiMonthName(new Date(opt.monthStr).getMonth())} {new Date(opt.monthStr).getFullYear() + 543} - วิชา {opt.subject}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                      
                      <div className="flex flex-col md:flex-row gap-4 items-end w-full">
                        <div className="flex-1 w-full">
                          <label className="block text-xs font-semibold text-slate-500 mb-2 uppercase tracking-wider flex items-center gap-1"><Calendar size={14}/> เลือกวันที่ในสัปดาห์</label>
                          <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} className="w-full px-4 py-3 bg-white/60 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all text-sm font-medium text-slate-700" />
                        </div>
                        {filteredSubjectList.length > 0 ? (
                          <div className="flex-1 w-full">
                            <label className="block text-xs font-semibold text-slate-500 mb-2 uppercase tracking-wider flex items-center gap-1"><BookOpen size={14}/> วิชา</label>
                            <select value={formData.subject} onChange={(e) => setFormData({...formData, subject: e.target.value})} className="w-full px-4 py-3 bg-white/60 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm font-medium text-slate-700">
                              <option value="" disabled>เลือกวิชา...</option>
                              {filteredSubjectList.map(sub => <option key={sub} value={sub}>{sub}</option>)}
                            </select>
                          </div>
                        ) : (<div className="flex-1 w-full text-sm text-slate-500 bg-white/40 p-3 rounded-xl border border-dashed text-center">ไม่มีข้อมูลวิชา</div>)}
                      </div>
                    </div>

                    {!formData.subject ? (
                      /* No subject selected — show prompt + timetable */
                      <div className="space-y-8 animate-fade-in-up">
                        <div className="glass-card rounded-[2rem] p-4 text-center flex flex-col items-center justify-center border-dashed border-2 border-indigo-200/50 bg-indigo-50/20">
                          <BookOpen size={32} className="text-indigo-500 mb-2 animate-pulse" />
                          <h3 className="text-md font-bold text-indigo-900">กรุณาเลือกวิชาในกล่องด้านบน หรือคลิกจากตารางเรียนด้านล่าง</h3>
                          <p className="text-xs text-indigo-700/70 mt-0.5">คลิกเพื่อคำนวณวันและรายวิชาที่สัมพันธ์ให้อัตโนมัติ</p>
                        </div>
                        <TimetableGrid subjectList={subjectList} onCellClick={handleTimetableCellClick} />
                      </div>
                    ) : isClubSubject(formData.subject) && activeClassStudents.length === 0 ? (
                      /* ชุมนุม — needs student selection first */
                      <div className="glass-card rounded-[2rem] p-10 flex flex-col items-center justify-center gap-5 border-2 border-dashed border-orange-200 bg-orange-50/30 animate-fade-in-up">
                        <div className="w-16 h-16 rounded-2xl bg-orange-100 flex items-center justify-center text-orange-600"><Users size={32} /></div>
                        <div className="text-center">
                          <h3 className="text-lg font-bold text-orange-900">กิจกรรมชุมนุม — เลือกรายชื่อนักเรียน</h3>
                          <p className="text-sm text-orange-700 mt-1">นักเรียนชุมนุมมาจากหลายห้อง กรุณาเลือกรายชื่อนักเรียนที่เข้าร่วม</p>
                        </div>
                        <button type="button" onClick={() => setShowStudentPicker(true)} className="bg-orange-500 text-white px-8 py-3 rounded-xl font-bold text-sm hover:bg-orange-600 transition-all shadow-lg shadow-orange-200 flex items-center gap-2 hover:scale-[1.02]">
                          <Users size={18} /> เลือกรายชื่อนักเรียน
                        </button>
                      </div>
                    ) : activeClassStudents.length > 0 && (
                      /* Normal class, ลูกเสือ, or ชุมนุม with selected students — show attendance table */
                      <div className="space-y-4 animate-fade-in-up">
                        {/* Cross-class info banners */}
                        {isScoutSubject(formData.subject) && (
                          <div className="glass-card rounded-2xl px-6 py-4 flex items-center gap-4 bg-emerald-50/50 border border-emerald-200">
                            <div className="w-10 h-10 rounded-xl bg-emerald-100 flex items-center justify-center text-emerald-600 shrink-0"><Users size={20} /></div>
                            <div className="flex-1">
                              <p className="font-bold text-emerald-800 text-sm">🪖 ลูกเสือ — โหลดรายชื่ออัตโนมัติ</p>
                              <p className="text-xs text-emerald-700">นักเรียนชั้น ป.4 - ม.3 ทั้งหมด {activeClassStudents.length} คน ({SCOUT_CLASSES.join(', ')})</p>
                            </div>
                          </div>
                        )}
                        {isClubSubject(formData.subject) && (
                          <div className="glass-card rounded-2xl px-6 py-4 flex items-center gap-4 bg-orange-50/50 border border-orange-200">
                            <div className="w-10 h-10 rounded-xl bg-orange-100 flex items-center justify-center text-orange-600 shrink-0"><Users size={20} /></div>
                            <div className="flex-1">
                              <p className="font-bold text-orange-800 text-sm">🎯 ชุมนุม — รายชื่อที่เลือก</p>
                              <p className="text-xs text-orange-700">{activeClassStudents.length} คนที่เลือก</p>
                            </div>
                            <button type="button" onClick={() => setShowStudentPicker(true)} className="px-4 py-2 bg-orange-100 text-orange-700 border border-orange-200 rounded-xl text-xs font-bold hover:bg-orange-200 transition-all">แก้ไขรายชื่อ</button>
                          </div>
                        )}

                        {/* Attendance table card */}
                        <div className="glass-card rounded-[2rem] overflow-hidden shadow-lg shadow-slate-200/50">
                          <div className="p-6 md:p-8 border-b border-slate-100/50 bg-white/40 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                            <div>
                              <h2 className="text-xl font-bold text-slate-800">
                                {isCrossClassSubject(formData.subject) ? `เช็คชื่อ — ${formData.subject}` : `เช็คชื่อเข้าเรียนห้อง ${globalClass}`}
                              </h2>
                              <p className="text-xs font-bold text-slate-500 mt-1 flex flex-wrap items-center gap-1.5">
                                <span>📅 ประจำเดือน:</span>
                                <span className="text-indigo-600 bg-indigo-50 px-2.5 py-0.5 rounded font-extrabold">{thaiMonthName(activeMonthIndex)} {activeYear + 543}</span>
                                <span className="text-slate-300">|</span>
                                <span>จำนวนวันเรียนทั้งหมด:</span>
                                <span className="text-slate-700 font-extrabold">{activeGridDates.filter(d => !excludedDates.includes(d)).length} วันทำการ</span>
                              </p>
                            </div>
                            <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto justify-end">
                              <div className="flex gap-2 items-center bg-white/60 px-3 py-1.5 rounded-xl border border-slate-100 text-xs text-slate-500 shadow-sm">
                                <span className="font-semibold text-slate-600">คำอธิบาย:</span>
                                <span className="flex items-center gap-1"><span className="w-4.5 h-4.5 rounded-md bg-emerald-500 inline-flex items-center justify-center text-white text-[10px] font-bold">✓</span> มา</span>
                                <span className="flex items-center gap-1"><span className="w-4.5 h-4.5 rounded-md bg-rose-500 inline-flex items-center justify-center text-white text-[10px] font-bold">✗</span> ขาด</span>
                                <span className="flex items-center gap-1"><span className="w-4.5 h-4.5 rounded-md bg-amber-500 inline-flex items-center justify-center text-white text-[10px] font-bold">⏰</span> สาย</span>
                              </div>
                              <div className="flex gap-2">
                                <button onClick={exportAttendanceExcel} className="bg-emerald-50 text-emerald-600 px-4 py-2.5 rounded-xl hover:bg-emerald-100 font-bold text-sm shadow-sm border border-emerald-100 flex items-center gap-2 transition-all">
                                  <Download size={18} /> ส่งออก Excel
                                </button>
                                <button onClick={handleSaveBulkAttendance} className="bg-indigo-600 text-white px-6 py-2.5 rounded-xl hover:bg-indigo-700 font-bold flex items-center gap-2 shadow-sm transition-all hover:scale-[1.02]">
                                  <Save size={18} /> บันทึกเดือนนี้
                                </button>
                              </div>
                            </div>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse">
                              <thead>
                                <tr className="bg-slate-50/50 text-xs uppercase tracking-wider text-slate-400">
                                  <th className="px-6 py-4 font-bold">เลขที่</th>
                                  <th className="px-6 py-4 font-bold">ชื่อ</th>
                                  <th className="px-6 py-4 font-bold">นามสกุล</th>
                                  
                                  {activeGridDates.map((dateStr, dIdx) => {
                                    const [yr, mo, dy] = dateStr.split('-').map(Number);
                                    const dateObj = new Date(yr, mo - 1, dy); // local time — ป้องกัน UTC offset ทำให้วันผิด
                                    const dayNames = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
                                    const dayName = dayNames[dateObj.getDay()];
                                    const isExcluded = excludedDates.includes(dateStr);
                                    
                                    const dayColors = {
                                      0: 'bg-rose-100/60 text-rose-800 border-b border-rose-200',      // อาทิตย์
                                      1: 'bg-yellow-100/60 text-yellow-800 border-b border-yellow-200', // จันทร์
                                      2: 'bg-pink-100/60 text-pink-800 border-b border-pink-200',       // อังคาร
                                      3: 'bg-emerald-100/60 text-emerald-800 border-b border-emerald-200', // พุธ
                                      4: 'bg-purple-100/60 text-purple-800 border-b border-purple-200', // พฤหัสบดี
                                      5: 'bg-cyan-100/60 text-cyan-800 border-b border-cyan-200',       // ศุกร์
                                      6: 'bg-slate-100/60 text-slate-800 border-b border-slate-200',    // เสาร์
                                    };
                                    let colorClass = dayColors[dateObj.getDay()] || 'bg-slate-100/60 text-slate-800 border-b border-slate-200';
                                    if (isExcluded) {
                                      colorClass = 'bg-slate-100/30 text-slate-400 border-b border-slate-200';
                                    }
                                  
                                  return (
                                    <th key={dIdx} className={`px-4 py-3 text-center min-w-[120px] transition-all ${isExcluded ? 'opacity-40' : ''} ${colorClass}`}>
                                      <div className="flex flex-col items-center gap-1">
                                        <label className="flex items-center gap-1 cursor-pointer select-none mb-1 text-[10px] text-slate-500 font-bold hover:text-slate-700">
                                          <input 
                                            type="checkbox" 
                                            checked={!isExcluded} 
                                            onChange={() => toggleDateExclusion(dateStr)} 
                                            className="w-3.5 h-3.5 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500/20 cursor-pointer"
                                          />
                                          เช็คชื่อ
                                        </label>
                                        <span className="font-extrabold text-sm">{dayName}</span>
                                        <span className="text-[10px] opacity-75 font-semibold">{formatHeaderDate(dateStr)}</span>
                                        <button 
                                          onClick={() => !isExcluded && handleMarkAllDayPresent(dateStr)} 
                                          disabled={isExcluded}
                                          className={`mt-1 px-2 py-0.5 rounded text-[10px] font-bold transition-all shadow-3xs ${
                                            isExcluded 
                                              ? 'bg-slate-100 text-slate-300 border border-slate-200 cursor-not-allowed' 
                                              : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-300'
                                          }`}
                                        >
                                          มาทุกคน
                                        </button>
                                      </div>
                                    </th>
                                  );
                                })}
                                
                                <th className="px-6 py-4 text-center bg-amber-200 text-amber-900 border-b border-amber-300 font-bold min-w-[80px]">รวม ({activeGridDates.filter(d => !excludedDates.includes(d)).length} วัน)</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100/50">
                              {activeClassStudents.map((student, idx) => {
                                const { firstName, lastName } = splitName(student.name);
                                
                                let totalPresent = 0;
                                const includedDates = activeGridDates.filter(d => !excludedDates.includes(d));
                                includedDates.forEach(dateStr => {
                                  const status = getAttendanceStatus(student.name, dateStr);
                                  if (status === 'present' || status === 'late') {
                                    totalPresent += 1;
                                  }
                                });

                                return (
                                  <tr key={idx} className="hover:bg-white/40">
                                    <td className="px-6 py-4 text-sm text-slate-500 font-semibold">{student.studentId}</td>
                                    <td className="px-6 py-4 text-sm font-semibold text-slate-700">{firstName}</td>
                                    <td className="px-6 py-4 text-sm font-semibold text-slate-700">{lastName}</td>
                                    {activeGridDates.map((dateStr, dIdx) => {
                                      const isExcluded = excludedDates.includes(dateStr);
                                      const status = getAttendanceStatus(student.name, dateStr);
                                      return (
                                        <td key={dIdx} className={`px-4 py-4 text-center transition-all ${isExcluded ? 'opacity-30' : ''}`}>
                                          <button 
                                            type="button"
                                            onClick={() => !isExcluded && toggleAttendanceStatus(student.name, dateStr)}
                                            disabled={isExcluded}
                                            className={`w-10 h-10 rounded-xl flex items-center justify-center mx-auto transition-all transform border ${
                                              isExcluded
                                                ? 'bg-slate-100 border-slate-200/60 text-slate-300 cursor-not-allowed shadow-none'
                                                : `hover:scale-105 active:scale-95 shadow-2xs ${
                                                    status === 'present' 
                                                      ? 'bg-emerald-500 border-emerald-600 text-white shadow-emerald-100' 
                                                      : status === 'absent' 
                                                        ? 'bg-rose-500 border-rose-600 text-white shadow-rose-100' 
                                                        : 'bg-amber-500 border-amber-600 text-white shadow-amber-100'
                                                  }`
                                            }`}
                                          >
                                            {isExcluded ? (
                                              <span className="text-xs font-bold text-slate-300">-</span>
                                            ) : (
                                              <>
                                                {status === 'present' && <CheckCircle size={18} />}
                                                {status === 'absent' && <AlertCircle size={18} />}
                                                {status === 'late' && <Clock size={18} />}
                                              </>
                                            )}
                                          </button>
                                        </td>
                                      );
                                    })}
                                    <td className="px-6 py-4 text-center font-extrabold text-sm text-amber-900 bg-amber-500/10 border-l border-slate-100">
                                      {totalPresent}
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                       </div>
                    )}
                  </>

                )}
              </div>
            )}

            {/* ---------------- ASSIGNMENTS TAB (TEACHER ONLY) ---------------- */}
            {activeTab === 'assignments' && isTeacher && (
              <div className="space-y-8">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {getAssignmentStats().map((stat) => (
                    <div key={stat.name} className="glass-card rounded-[2rem] p-6 hover:-translate-y-1 transition-all">
                      <div className="flex items-center gap-4">
                        <div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${stat.bg}`}>{stat.icon}</div>
                        <div>
                          <p className="text-slate-500 text-sm font-medium mb-1">{stat.name}</p>
                          <h3 className={`text-3xl font-bold ${stat.text}`}>{stat.value} <span className="text-sm font-normal text-slate-400">รายการ</span></h3>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>

                {globalClass === 'all' ? (
                  <div className="glass-card rounded-[2rem] p-12 text-center flex flex-col items-center justify-center border-dashed border-2 border-slate-300/50 bg-white/40">
                    <BookOpen size={48} className="text-purple-300 mb-4" />
                    <h3 className="text-lg font-bold text-slate-700">กรุณาเลือกชั้นเรียนที่แถบด้านบน</h3>
                  </div>
                ) : (
                  <>
                    <div className="glass-card rounded-[2rem] p-6 flex flex-col gap-4 shadow-sm">
                      {pastAssignmentOptions.length > 0 && (
                        <div className="w-full pb-4 border-b border-slate-100">
                          <label className="block text-xs font-semibold text-purple-500 mb-2 uppercase tracking-wider flex items-center gap-1"><Edit3 size={14}/> เลือกงานเดิมเพื่อแก้ไข</label>
                          <select onChange={(e) => handleLoadPastAssignment(e.target.value)} className="w-full px-4 py-3 bg-purple-50/50 border border-purple-100 rounded-xl focus:ring-2 focus:ring-purple-500/20 outline-none text-sm font-medium text-purple-700 cursor-pointer transition-colors hover:bg-purple-50">
                            <option value="">-- สร้างงานใหม่ --</option>
                            {pastAssignmentOptions.map((opt, i) => (
                              <option key={i} value={`${opt.assignment}|${opt.subject}|${opt.dueDate}`}>
                                งาน: {opt.assignment} (วิชา: {opt.subject}) - กำหนดส่ง {opt.dueDate}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                      
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 w-full">
                        <div>
                          <label className="block text-xs font-semibold text-slate-500 mb-2">ชื่องาน</label>
                          <input type="text" value={assignForm.name} onChange={(e) => setAssignForm({...assignForm, name: e.target.value})} className="w-full px-4 py-3 bg-white/60 border rounded-xl" />
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-slate-500 mb-2">กำหนดส่ง</label>
                          <input type="date" value={assignForm.dueDate} onChange={(e) => setAssignForm({...assignForm, dueDate: e.target.value})} className="w-full px-4 py-3 bg-white/60 border rounded-xl" />
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-slate-500 mb-2">วิชา</label>
                          <select value={assignForm.subject} onChange={(e) => setAssignForm({...assignForm, subject: e.target.value})} className="w-full px-4 py-3 bg-white/60 border rounded-xl">
                            <option value="" disabled>เลือกวิชา...</option>
                            {filteredSubjectList.map(sub => <option key={sub} value={sub}>{sub}</option>)}
                          </select>
                        </div>
                      </div>
                    </div>

                    {activeClassStudents.length > 0 && (
                      <div className="glass-card rounded-[2rem] overflow-hidden shadow-lg shadow-slate-200/50">
                        <div className="p-6 md:p-8 border-b border-slate-100/50 bg-white/40 flex justify-between items-center">
                          <h2 className="text-xl font-bold text-slate-800">เช็คส่งงานห้อง {globalClass}</h2>
                          <div className="flex gap-3">
                            <button onClick={exportAssignmentExcel} className="bg-emerald-50 text-emerald-600 px-4 py-2.5 rounded-xl hover:bg-emerald-100 font-bold text-sm shadow-sm border border-emerald-100 flex items-center gap-2 transition-all">
                              <Download size={18} /> ส่งออก Excel
                            </button>
                            <button onClick={handleSaveBulkAssignment} className="bg-purple-600 text-white px-6 py-2.5 rounded-xl hover:bg-purple-700 font-medium flex items-center gap-2">
                              <Save size={18} /> บันทึก
                            </button>
                          </div>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-50/50 text-xs uppercase tracking-wider text-slate-400">
                                <th className="px-6 py-4">เลขที่</th>
                                <th className="px-6 py-4">ชื่อ</th>
                                <th className="px-6 py-4 text-center">สถานะ</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100/50">
                              {activeClassStudents.map((student, idx) => (
                                <tr key={idx} className="hover:bg-white/40">
                                  <td className="px-6 py-4 text-sm text-slate-500">{student.studentId}</td>
                                  <td className="px-6 py-4 text-sm font-medium text-slate-700">{student.name}</td>
                                  <td className="px-6 py-4">
                                    <div className="flex justify-center gap-2">
                                      {['completed', 'pending', 'overdue'].map((s) => (
                                        <label key={s} className={`cursor-pointer px-4 py-2 rounded-full text-xs font-semibold border ${getAssignmentStatus(student.name) === s ? (s === 'completed' ? 'bg-emerald-500 text-white border-emerald-500' : s === 'pending' ? 'bg-blue-500 text-white border-blue-500' : 'bg-rose-500 text-white border-rose-500') : 'bg-white text-slate-500'}`}>
                                          <input type="radio" className="hidden" checked={getAssignmentStatus(student.name) === s} onChange={() => handleAssignStatusChange(student.name, s)} />
                                          {s === 'completed' ? 'ส่งแล้ว' : s === 'pending' ? 'กำลังทำ' : 'เลยกำหนด'}
                                        </label>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </>
                )}
              </div>
            )}

            {/* ---------------- STUDENT SUMMARY TAB (PUBLIC) ---------------- */}
            {activeTab === 'summary' && (
              <div className="space-y-6 animate-fade-in-up">
                
                {/* Search Filters */}
                <div className="glass-card rounded-[2rem] p-8 flex flex-col md:flex-row gap-6 relative overflow-hidden">
                  <div className="absolute right-0 top-0 opacity-5 w-64 h-64 -translate-y-1/2 translate-x-1/4">
                    <Award size={250} />
                  </div>
                  <div className="relative z-10 w-full flex flex-col md:flex-row gap-6 items-end">
                    <div className="flex-1 w-full">
                      <label className="block text-sm font-bold text-slate-600 mb-2">1. เลือกชั้นเรียน</label>
                      <select value={summaryClass} onChange={(e) => {setSummaryClass(e.target.value); setSummaryStudent('');}} className="w-full px-5 py-3.5 bg-white border border-slate-200 rounded-2xl focus:ring-4 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none text-slate-700 font-medium shadow-sm transition-all">
                        <option value="" disabled>-- กรุณาเลือกชั้นเรียน --</option>
                        {classList.map(cls => <option key={cls} value={cls}>ห้อง {cls}</option>)}
                      </select>
                    </div>
                    
                    <div className="flex-1 w-full">
                      <label className="block text-sm font-bold text-slate-600 mb-2">2. เลือกชื่อนักเรียน</label>
                      <select value={summaryStudent} onChange={(e) => setSummaryStudent(e.target.value)} disabled={!summaryClass} className="w-full px-5 py-3.5 bg-white border border-slate-200 rounded-2xl focus:ring-4 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none text-slate-700 font-medium shadow-sm transition-all disabled:opacity-50 disabled:bg-slate-50">
                        <option value="" disabled>-- กรุณาเลือกชื่อนักเรียน --</option>
                        {summaryStudentList.map(name => <option key={name} value={name}>{name}</option>)}
                      </select>
                    </div>
                  </div>
                </div>

                {!summaryStudent ? (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 animate-fade-in-up">
                    <div className="glass-card rounded-[2rem] p-8 text-center bg-white/40 flex flex-col items-center shadow-sm">
                      <div className="p-3 bg-emerald-100 text-emerald-600 rounded-2xl mb-4"><Calendar size={28} /></div>
                      <h3 className="text-xl font-extrabold text-slate-700 mb-2">สถิติการเข้าเรียน</h3>
                      <p className="text-slate-500 text-sm mb-6">{summaryClass ? `เฉพาะนักเรียนห้อง ${summaryClass}` : 'นักเรียนทั้งหมดทุกระดับชั้น'}</p>
                      
                      {summaryAttStats.length > 0 ? (
                        <div className="h-64 w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie data={summaryAttStats} cx="50%" cy="50%" innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value" stroke="none">
                                {summaryAttStats.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.fill} />)}
                              </Pie>
                              <Tooltip formatter={(value) => [`${value} ครั้ง`, 'จำนวน']} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                              <Legend verticalAlign="bottom" height={36} iconType="circle" />
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                      ) : (
                        <div className="h-64 flex items-center justify-center text-slate-400 bg-slate-50 w-full rounded-2xl border border-dashed">ไม่มีข้อมูลการเข้าเรียน</div>
                      )}
                    </div>
                    
                    <div className="glass-card rounded-[2rem] p-8 text-center bg-white/40 flex flex-col items-center shadow-sm">
                      <div className="p-3 bg-blue-100 text-blue-600 rounded-2xl mb-4"><BookOpen size={28} /></div>
                      <h3 className="text-xl font-extrabold text-slate-700 mb-2">สถิติการส่งงาน</h3>
                      <p className="text-slate-500 text-sm mb-6">{summaryClass ? `เฉพาะนักเรียนห้อง ${summaryClass}` : 'นักเรียนทั้งหมดทุกระดับชั้น'}</p>
                      
                      {summaryAssignStats.length > 0 ? (
                        <div className="h-64 w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie data={summaryAssignStats} cx="50%" cy="50%" innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value" stroke="none">
                                {summaryAssignStats.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.fill} />)}
                              </Pie>
                              <Tooltip formatter={(value) => [`${value} ชิ้น`, 'จำนวน']} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                              <Legend verticalAlign="bottom" height={36} iconType="circle" />
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                      ) : (
                        <div className="h-64 flex items-center justify-center text-slate-400 bg-slate-50 w-full rounded-2xl border border-dashed">ไม่มีข้อมูลการส่งงาน</div>
                      )}
                    </div>
                  </div>
                ) : (
                  <div className="space-y-6 animate-fade-in-up">
                    
                    {/* Student Profile Card */}
                    <div className="glass-card rounded-[2rem] p-8 flex flex-col md:flex-row items-center gap-8 bg-gradient-to-br from-white to-emerald-50/50 border-emerald-100/50 shadow-lg shadow-emerald-500/5">
                      <div className="w-24 h-24 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-600 shadow-inner">
                        <User size={48} />
                      </div>
                      <div className="text-center md:text-left flex-1">
                        <h2 className="text-3xl font-extrabold text-slate-800 mb-2">{summaryStudent}</h2>
                        <div className="flex flex-wrap items-center justify-center md:justify-start gap-3">
                          <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-white rounded-full text-sm font-semibold text-slate-600 border border-slate-200 shadow-sm"><Users size={14}/> ห้อง {summaryClass}</span>
                          <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-white rounded-full text-sm font-semibold text-slate-600 border border-slate-200 shadow-sm"><MapPin size={14}/> รหัสนักเรียน: {normalizedStudents.find(s => s.name === summaryStudent)?.studentId || '-'}</span>
                        </div>
                      </div>
                      
                      {/* Mini Stats inside Profile */}
                      <div className="flex gap-4">
                        <div className="text-center px-6 py-4 bg-white rounded-2xl shadow-sm border border-slate-100">
                          <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">การเข้าเรียน</p>
                          <p className="text-2xl font-black text-emerald-500">{studentAttHistory.filter(h => h.status === 'present').length} <span className="text-sm font-semibold text-slate-400">ครั้ง</span></p>
                        </div>
                        <div className="text-center px-6 py-4 bg-white rounded-2xl shadow-sm border border-slate-100">
                          <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">งานที่ส่งแล้ว</p>
                          <p className="text-2xl font-black text-emerald-500">{studentAssignHistory.filter(h => h.status === 'completed').length} <span className="text-sm font-semibold text-slate-400">ชิ้น</span></p>
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                      {/* Attendance History List */}
                      <div className="glass-card rounded-[2rem] p-8">
                        <div className="flex items-center gap-3 mb-6">
                          <div className="p-2 bg-indigo-100 text-indigo-600 rounded-xl"><Calendar size={20} /></div>
                          <h3 className="text-lg font-bold text-slate-800">ประวัติการเข้าเรียน</h3>
                        </div>
                        <div className="space-y-4 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                          {studentAttHistory.length > 0 ? studentAttHistory.map((h, i) => (
                            <div key={i} className="flex items-center justify-between p-4 rounded-2xl bg-white/60 border border-slate-100 hover:shadow-md transition-shadow">
                              <div>
                                <p className="font-semibold text-slate-700">{h.subject}</p>
                                <p className="text-xs text-slate-500 mt-1">{h.date}</p>
                              </div>
                              <span className={`px-3 py-1.5 rounded-full text-xs font-bold ${
                                h.status === 'present' ? 'bg-emerald-100 text-emerald-700' :
                                h.status === 'absent' ? 'bg-rose-100 text-rose-700' : 'bg-amber-100 text-amber-700'
                              }`}>
                                {h.status === 'present' ? 'มาเรียน' : h.status === 'absent' ? 'ขาดเรียน' : 'มาสาย'}
                              </span>
                            </div>
                          )) : <p className="text-center text-slate-400 py-10">ยังไม่มีข้อมูลการเข้าเรียน</p>}
                        </div>
                      </div>

                      {/* Assignment History List */}
                      <div className="glass-card rounded-[2rem] p-8">
                        <div className="flex items-center gap-3 mb-6">
                          <div className="p-2 bg-purple-100 text-purple-600 rounded-xl"><BookOpen size={20} /></div>
                          <h3 className="text-lg font-bold text-slate-800">ประวัติการส่งงาน</h3>
                        </div>
                        <div className="space-y-4 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                          {studentAssignHistory.length > 0 ? studentAssignHistory.map((h, i) => (
                            <div key={i} className="flex items-center justify-between p-4 rounded-2xl bg-white/60 border border-slate-100 hover:shadow-md transition-shadow">
                              <div>
                                <p className="font-semibold text-slate-700">{h.assignment}</p>
                                <p className="text-xs text-slate-500 mt-1">{h.subject} • 📅 {h.dueDate}</p>
                              </div>
                              <span className={`px-3 py-1.5 rounded-full text-xs font-bold ${
                                h.status === 'completed' ? 'bg-emerald-100 text-emerald-700' :
                                h.status === 'overdue' ? 'bg-rose-100 text-rose-700' : 'bg-blue-100 text-blue-700'
                              }`}>
                                {h.status === 'completed' ? 'ส่งแล้ว' : h.status === 'overdue' ? 'เลยกำหนด' : 'กำลังทำ'}
                              </span>
                            </div>
                          )) : <p className="text-center text-slate-400 py-10">ยังไม่มีข้อมูลการส่งงาน</p>}
                        </div>
                      </div>
                    </div>

                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </main>

      {timetableAlert.isOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-fade-in">
          <div className="bg-white rounded-[2.5rem] p-8 max-w-md w-full shadow-2xl border border-slate-100 flex flex-col gap-6 relative animate-scale-up">
            <button 
              onClick={() => setTimetableAlert({ isOpen: false, subject: '', className: '' })}
              className="absolute top-6 right-6 w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-500 hover:bg-slate-200 transition-colors font-bold"
            >
              ✕
            </button>
            
            <div className="flex items-center gap-4 border-b border-slate-100 pb-4">
              <div className="w-12 h-12 rounded-2xl bg-amber-100 flex items-center justify-center text-amber-600">
                <BookOpen size={24} />
              </div>
              <div>
                <h3 className="text-xl font-bold text-slate-800">ไม่พบวิชาใน Google Sheets</h3>
                <p className="text-sm text-slate-500">วิชา: {timetableAlert.subject} ห้อง {timetableAlert.className === 'all' ? 'ทุกชั้นเรียน' : timetableAlert.className}</p>
              </div>
            </div>
            
            <div className="space-y-4 text-sm text-slate-600">
              <p className="font-semibold text-slate-700">เนื่องจากระบบนี้ดึงข้อมูลจาก Google Sheets ของคุณแบบเรียลไทม์ คุณสามารถเพิ่มวิชานี้ได้ตามขั้นตอนดังนี้:</p>
              <ol className="list-decimal pl-5 space-y-2.5">
                <li>เปิด Google Sheets ของคุณ</li>
                <li>ไปที่แผ่นงานชื่อ <span className="px-2 py-0.5 bg-indigo-50 text-indigo-600 font-mono rounded font-bold">Subjects</span></li>
                <li>พิมพ์เพิ่มรายชื่อวิชาใหม่ในแถวล่างสุดเป็น: <span className="px-2 py-0.5 bg-emerald-50 text-emerald-700 font-mono rounded font-bold">{timetableAlert.subject === 'ซ่อมเสริม(ดนตรี)' ? `ซ่อมเสริม(ดนตรี) ${timetableAlert.className}` : `${timetableAlert.subject} ${timetableAlert.className}`}</span></li>
                {timetableAlert.className !== 'all' && (
                  <li>
                    ตรวจสอบว่าห้องเรียน <span className="px-2 py-0.5 bg-amber-50 text-amber-700 font-mono rounded font-bold">{timetableAlert.className}</span> มีนักเรียนในแผ่นงาน <span className="px-2 py-0.5 bg-indigo-50 text-indigo-600 font-mono rounded font-bold">Students</span> แล้ว
                  </li>
                )}
              </ol>
              <div className="bg-indigo-50 border border-indigo-100 rounded-2xl p-4 text-xs text-indigo-700 leading-relaxed mt-2 flex gap-3">
                <Sparkles size={24} className="shrink-0 text-indigo-500 animate-pulse" />
                <div>
                  หลังจากกรอกใน Google Sheets แล้ว ให้กดปุ่ม <strong>"โหลดข้อมูลใหม่"</strong> บนแผงควบคุมระบบ (หรือรีเฟรชหน้าเว็บ) เพื่อดึงข้อมูลมาใช้งานได้ทันที!
                </div>
              </div>
            </div>
            
            <div className="flex gap-3 justify-end pt-4 border-t border-slate-100">
              <button 
                onClick={() => setTimetableAlert({ isOpen: false, subject: '', className: '' })}
                className="px-6 py-2.5 bg-slate-100 text-slate-700 hover:bg-slate-200 rounded-xl font-semibold transition-all text-sm animate-scale-up"
              >
                เข้าใจแล้ว
              </button>
              {sheetUrl && (
                <a 
                  href={sheetUrl} 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="px-6 py-2.5 bg-indigo-600 text-white hover:bg-indigo-700 rounded-xl font-semibold transition-all text-sm flex items-center gap-2"
                >
                  เปิด Google Sheets <ArrowRight size={16} />
                </a>
              )}
            </div>
          </div>
        </div>
      )}
      <style dangerouslySetInnerHTML={{__html: `
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 20px; }
        .custom-scrollbar:hover::-webkit-scrollbar-thumb { background-color: #94a3b8; }
      `}} />
    </div>
  );
}
