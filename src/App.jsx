import React, { useState, useEffect, useMemo } from 'react';
import { BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { CheckCircle, AlertCircle, BookOpen, Users, Calendar, Settings, ArrowRight, LayoutDashboard, ChevronRight, Clock, MapPin, Sparkles, Filter, Save, Edit3, Search, Lock, User, LogOut, Award, Download } from 'lucide-react';
import * as XLSX from 'xlsx';

const normalizeRow = (row) => {
  return {
    name: row.Name || row['ชื่อนักเรียน'] || row['ชื่อ'] || row['ชื่อ-สกุล'] || row['ชื่อ-นามสกุล'] || row['ชื่อ - สกุล'] || '-',
    subject: row.Subject || row['วิชา'] || row['รายวิชา'] || '-',
    date: row.Date || row['วันที่'] || row['วันที่เช็คชื่อ'] || '-',
    status: row.Status || row['สถานะ'] || '-',
    assignment: row.Assignment || row['ชื่องาน'] || row['งาน'] || '-',
    dueDate: row.DueDate || row['กำหนดส่ง'] || row['วันส่ง'] || '-',
    studentId: row.StudentID || row['เลขที่'] || row['รหัส'] || row['รหัสนักเรียน'] || '-',
    className: row.Class || row['ชั้นเรียน'] || row['ชั้น'] || row['ห้อง'] || '-'
  };
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
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [formData, setFormData] = useState({ subject: '' });
  const [attendanceList, setAttendanceList] = useState({});

  // ---------------- ASSIGNMENT STATES ----------------
  const [assignForm, setAssignForm] = useState({ subject: '', name: '', dueDate: new Date().toISOString().split('T')[0] });
  const [assignStatusList, setAssignStatusList] = useState({});

  // ---------------- SUMMARY STATES ----------------
  const [summaryClass, setSummaryClass] = useState('');
  const [summaryStudent, setSummaryStudent] = useState('');
  
  const [animateTab, setAnimateTab] = useState(false);

  useEffect(() => {
    setAnimateTab(true);
    const timer = setTimeout(() => setAnimateTab(false), 600);
    return () => clearTimeout(timer);
  }, [activeTab, user]);

  const extractSheetId = (url) => {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
  };

  const fetchSheetData = async (sheetId, sheetName, silent = false) => {
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
  };

  // Fetch data on initial load
  useEffect(() => {
    if (sheetUrl && !showSetup) {
      const sheetId = extractSheetId(sheetUrl);
      if (sheetId) {
        loadAllData(sheetId);
      }
    }
  }, [sheetUrl, showSetup]);

  const loadAllData = async (sheetId) => {
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
  };

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
    } catch (error) {
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
      .filter(s => s.name && s.name !== '-' && s.name.trim() !== '');
  }, [studentsData]);
  
  const classList = useMemo(() => {
    return [...new Set(normalizedStudents.map(s => s.className).filter(c => c !== '-'))].sort();
  }, [normalizedStudents]);

  const subjectList = useMemo(() => {
    return [...new Set(subjectsData.map(normalizeRow).map(s => s.subject).filter(s => s !== '-'))];
  }, [subjectsData]);

  // Active Class Students for Checklists
  const activeClassStudents = useMemo(() => {
    if (globalClass === 'all') return [];
    return normalizedStudents.filter(s => s.className === globalClass);
  }, [normalizedStudents, globalClass]);

  // Students list for Summary Dropdown
  const summaryStudentList = useMemo(() => {
    if (!summaryClass) return [];
    return normalizedStudents.filter(s => s.className === summaryClass).map(s => s.name).sort();
  }, [normalizedStudents, summaryClass]);

  // Init Checklists
  useEffect(() => {
    const initialAtt = {};
    const initialAssign = {};
    activeClassStudents.forEach(s => {
      initialAtt[s.name] = 'present';
      initialAssign[s.name] = 'pending';
    });
    setAttendanceList(initialAtt);
    setAssignStatusList(initialAssign);
  }, [activeClassStudents]);

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

  // Options for past attendance
  const pastAttendanceOptions = useMemo(() => {
    const options = [];
    const seen = new Set();
    filteredAttendanceHistory.forEach(record => {
      const key = `${record.date}|${record.subject}`;
      if (!seen.has(key) && record.date !== '-' && record.subject !== '-') {
        seen.add(key);
        options.push({ date: record.date, subject: record.subject });
      }
    });
    return options.sort((a, b) => new Date(b.date) - new Date(a.date));
  }, [filteredAttendanceHistory]);

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
  const handleStatusChange = (name, status) => setAttendanceList(prev => ({ ...prev, [name]: status }));
  const handleAssignStatusChange = (name, status) => setAssignStatusList(prev => ({ ...prev, [name]: status }));

  const handleLoadPastAttendance = (val) => {
    if (!val) {
      setSelectedDate(new Date().toISOString().split('T')[0]);
      setFormData({ subject: '' });
      const initialAtt = {};
      activeClassStudents.forEach(s => initialAtt[s.name] = 'present');
      setAttendanceList(initialAtt);
      return;
    }
    const [date, subject] = val.split('|');
    setSelectedDate(date);
    setFormData({ subject });
    
    const pastRecords = filteredAttendanceHistory.filter(r => r.date === date && r.subject === subject);
    const newAttList = {};
    activeClassStudents.forEach(s => {
      const record = pastRecords.find(r => r.name === s.name);
      newAttList[s.name] = record ? record.status : 'present';
    });
    setAttendanceList(newAttList);
  };

  const handleLoadPastAssignment = (val) => {
    if (!val) {
      setAssignForm({ subject: '', name: '', dueDate: new Date().toISOString().split('T')[0] });
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
    if (globalClass === 'all') return alert('กรุณาเลือกชั้นเรียนก่อน');
    if (!formData.subject) return alert('กรุณาเลือกวิชาก่อนบันทึก');
    
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
      const payloadData = activeClassStudents
        .filter(student => student.name && student.name !== '-' && student.name.trim() !== '')
        .map(student => ({
          Name: student.name,
          Subject: formData.subject,
          Date: selectedDate,
          Status: attendanceList[student.name] || 'present'
        }));

      if (payloadData.length === 0) {
        setIsSaving(false);
        return alert('❌ ไม่พบรายชื่อนักเรียนที่มีข้อมูลถูกต้องสำหรับบันทึก');
      }

      // Debug log
      console.log('Attempting to save to (no-cors mode):', scriptUrl.substring(0, 45) + '...');

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
      
      // เนื่องจากโหมด no-cors จะไม่คืนค่า success/error เราจึงต้องถือว่าส่งสำเร็จ
      alert('✅ ส่งข้อมูลการเข้าเรียนเรียบร้อยแล้ว!\n(โปรดตรวจสอบความถูกต้องใน Google Sheets อีกครั้ง)');
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
      const payloadData = activeClassStudents
        .filter(student => student.name && student.name !== '-' && student.name.trim() !== '')
        .map(student => ({
          Name: student.name,
          Subject: assignForm.subject,
          Assignment: assignForm.name,
          DueDate: assignForm.dueDate,
          Status: assignStatusList[student.name] || 'pending'
        }));

      if (payloadData.length === 0) {
        setIsSaving(false);
        return alert('❌ ไม่พบรายชื่อนักเรียนที่มีข้อมูลถูกต้องสำหรับบันทึก');
      }

      // Debug log
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


  const DecorativeBackground = () => (
    <div className="fixed inset-0 overflow-hidden pointer-events-none z-[-1] bg-slate-50">
      <div className="absolute top-[-10%] left-[-10%] w-96 h-96 bg-purple-300 rounded-full mix-blend-multiply filter blur-[80px] opacity-60 animate-blob"></div>
      <div className="absolute top-[20%] right-[-10%] w-[30rem] h-[30rem] bg-cyan-300 rounded-full mix-blend-multiply filter blur-[100px] opacity-60 animate-blob animation-delay-2000"></div>
      <div className="absolute bottom-[-20%] left-[20%] w-[40rem] h-[40rem] bg-pink-200 rounded-full mix-blend-multiply filter blur-[120px] opacity-60 animate-blob animation-delay-4000"></div>
    </div>
  );

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
              <select value={globalClass} onChange={(e) => setGlobalClass(e.target.value)} className="w-full sm:w-auto bg-transparent border-none text-indigo-700 font-bold text-sm focus:ring-0 outline-none cursor-pointer pr-4">
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
          <div className={animateTab ? 'opacity-0 translate-y-4' : 'opacity-100 translate-y-0 transition-all duration-500 ease-out'}>
            
            {/* ---------------- ATTENDANCE TAB (TEACHER ONLY) ---------------- */}
            {activeTab === 'attendance' && isTeacher && (
              <div className="space-y-8">
                {/* Stats */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {getAttendanceStats().map((stat, i) => (
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

                {globalClass === 'all' ? (
                  <div className="glass-card rounded-[2rem] p-12 text-center flex flex-col items-center justify-center border-dashed border-2 border-slate-300/50 bg-white/40">
                    <Users size={48} className="text-indigo-300 mb-4" />
                    <h3 className="text-lg font-bold text-slate-700">กรุณาเลือกชั้นเรียนที่แถบด้านบน</h3>
                    <p className="text-sm text-slate-500 mt-1">เพื่อเริ่มทำการเช็คชื่อประจำวัน</p>
                  </div>
                ) : (
                  <>
                    <div className="glass-card rounded-[2rem] p-6 flex flex-col gap-4 shadow-sm">
                      {pastAttendanceOptions.length > 0 && (
                        <div className="w-full pb-4 border-b border-slate-100">
                          <label className="block text-xs font-semibold text-indigo-500 mb-2 uppercase tracking-wider flex items-center gap-1"><Edit3 size={14}/> เลือกข้อมูลเดิมเพื่อแก้ไข</label>
                          <select onChange={(e) => handleLoadPastAttendance(e.target.value)} className="w-full px-4 py-3 bg-indigo-50/50 border border-indigo-100 rounded-xl focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm font-medium text-indigo-700 cursor-pointer transition-colors hover:bg-indigo-50">
                            <option value="">-- สร้างการเช็คชื่อใหม่ --</option>
                            {pastAttendanceOptions.map((opt, i) => (
                              <option key={i} value={`${opt.date}|${opt.subject}`}>
                                วันที่ {opt.date} - วิชา {opt.subject}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                      
                      <div className="flex flex-col md:flex-row gap-4 items-end w-full">
                        <div className="flex-1 w-full">
                          <label className="block text-xs font-semibold text-slate-500 mb-2 uppercase tracking-wider flex items-center gap-1"><Calendar size={14}/> วันที่เช็คชื่อ</label>
                          <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} className="w-full px-4 py-3 bg-white/60 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all text-sm font-medium text-slate-700" />
                        </div>
                        {subjectList.length > 0 ? (
                          <div className="flex-1 w-full">
                            <label className="block text-xs font-semibold text-slate-500 mb-2 uppercase tracking-wider flex items-center gap-1"><BookOpen size={14}/> วิชา</label>
                            <select value={formData.subject} onChange={(e) => setFormData({...formData, subject: e.target.value})} className="w-full px-4 py-3 bg-white/60 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm font-medium text-slate-700">
                              <option value="" disabled>เลือกวิชา...</option>
                              {subjectList.map(sub => <option key={sub} value={sub}>{sub}</option>)}
                            </select>
                          </div>
                        ) : (<div className="flex-1 w-full text-sm text-slate-500 bg-white/40 p-3 rounded-xl border border-dashed text-center">ไม่มีข้อมูลวิชา</div>)}
                      </div>
                    </div>

                    {activeClassStudents.length > 0 && (
                      <div className="glass-card rounded-[2rem] overflow-hidden shadow-lg shadow-slate-200/50">
                        <div className="p-6 md:p-8 border-b border-slate-100/50 bg-white/40 flex justify-between items-center">
                          <div>
                            <h2 className="text-xl font-bold text-slate-800">เช็คชื่อห้อง {globalClass}</h2>
                          </div>
                          <div className="flex gap-3">
                            <button onClick={exportAttendanceExcel} className="bg-emerald-50 text-emerald-600 px-4 py-2.5 rounded-xl hover:bg-emerald-100 font-bold text-sm shadow-sm border border-emerald-100 flex items-center gap-2 transition-all">
                              <Download size={18} /> ส่งออก Excel
                            </button>
                            <button onClick={handleSaveBulkAttendance} className="bg-indigo-600 text-white px-6 py-2.5 rounded-xl hover:bg-indigo-700 font-medium flex items-center gap-2">
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
                                      {['present', 'absent', 'late'].map((s) => (
                                        <label key={s} className={`cursor-pointer px-4 py-2 rounded-full text-xs font-semibold border ${attendanceList[student.name] === s ? (s === 'present' ? 'bg-emerald-500 text-white border-emerald-500' : s === 'absent' ? 'bg-rose-500 text-white border-rose-500' : 'bg-amber-500 text-white border-amber-500') : 'bg-white text-slate-500'}`}>
                                          <input type="radio" className="hidden" checked={attendanceList[student.name] === s} onChange={() => handleStatusChange(student.name, s)} />
                                          {s === 'present' ? 'มา' : s === 'absent' ? 'ขาด' : 'สาย'}
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

            {/* ---------------- ASSIGNMENTS TAB (TEACHER ONLY) ---------------- */}
            {activeTab === 'assignments' && isTeacher && (
              <div className="space-y-8">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {getAssignmentStats().map((stat, i) => (
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
                            {subjectList.map(sub => <option key={sub} value={sub}>{sub}</option>)}
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
                                        <label key={s} className={`cursor-pointer px-4 py-2 rounded-full text-xs font-semibold border ${assignStatusList[student.name] === s ? (s === 'completed' ? 'bg-emerald-500 text-white border-emerald-500' : s === 'pending' ? 'bg-blue-500 text-white border-blue-500' : 'bg-rose-500 text-white border-rose-500') : 'bg-white text-slate-500'}`}>
                                          <input type="radio" className="hidden" checked={assignStatusList[student.name] === s} onChange={() => handleAssignStatusChange(student.name, s)} />
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

      <style dangerouslySetInnerHTML={{__html: `
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 20px; }
        .custom-scrollbar:hover::-webkit-scrollbar-thumb { background-color: #94a3b8; }
      `}} />
    </div>
  );
}
