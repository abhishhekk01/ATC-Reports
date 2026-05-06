/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, FormEvent, useRef } from 'react';
import { 
  Plus, 
  Home, 
  Users, 
  BarChart3, 
  MoreVertical, 
  X, 
  Phone, 
  GraduationCap,
  UserPlus,
  ChevronRight,
  Search,
  LogOut,
  User as UserIcon,
  Loader2,
  Edit2,
  Trash2,
  FileSpreadsheet,
  ArrowLeft,
  CheckCircle2,
  AlertCircle,
  FileText
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { Student, TabType, DailyReport } from './types.ts';
import { 
  auth, 
  db, 
  loginWithGoogle, 
  logout, 
  handleFirestoreError, 
  OperationType 
} from './firebase';
import { 
  onAuthStateChanged, 
  User 
} from 'firebase/auth';
import { 
  collection, 
  onSnapshot, 
  query, 
  where, 
  addDoc, 
  orderBy,
  serverTimestamp,
  Timestamp,
  updateDoc,
  deleteDoc,
  doc
} from 'firebase/firestore';

const CLASSES = [
  'Nursery', 'LKG', 'UKG',
  '1st', '2nd', '3rd', '4th', '5th',
  '6th', '7th', '8th', '9th', '10th',
  '11th', '12th'
];

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [activeTab, setActiveTab] = useState<TabType>('home');
  const [students, setStudents] = useState<Student[]>([]);
  const [studentsLoading, setStudentsLoading] = useState(true);
  const [reports, setReports] = useState<DailyReport[]>([]);
  const [reportsLoading, setReportsLoading] = useState(true);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [isReportModalOpen, setIsReportModalOpen] = useState(false);
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  const [reportingStudent, setReportingStudent] = useState<Student | null>(null);
  const [historyStudent, setHistoryStudent] = useState<Student | null>(null);
  const [historyMonth, setHistoryMonth] = useState(new Date().toISOString().slice(0, 7));
  const [currentReport, setCurrentReport] = useState<DailyReport | null>(null);
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  // Form State - Student
  const [newName, setNewName] = useState('');
  const [newClass, setNewClass] = useState('1st');
  const [newShift, setNewShift] = useState<'1st' | '2nd' | '3rd'>('1st');
  const [newContact, setNewContact] = useState('');

  // Delete Confirmation State
  const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);
  const [isClearReportsConfirmOpen, setIsClearReportsConfirmOpen] = useState(false);
  const [studentToDelete, setStudentToDelete] = useState<Student | null>(null);
  const [reportViewMode, setReportViewMode] = useState<'day' | 'month'>('day');
  const [isReportsMoreMenuOpen, setIsReportsMoreMenuOpen] = useState(false);

  // Helper to check if item is in current shift
  const isInCurrentShift = (student: Student) => {
    if (activeTab === '1st') return student.shift === '1st';
    if (activeTab === '2nd') return student.shift === '2nd';
    if (activeTab === '3rd') return student.shift === '3rd';
    return true;
  };

  const getMonthlyStats = (studentId: string) => {
    const studentReports = reports.filter(r => 
      r.studentId === studentId && 
      r.date.startsWith(reportDate.slice(0, 7))
    );

    const presentDays = studentReports.filter(r => r.attendance).length;
    const goodFocus = studentReports.filter(r => r.attendance && r.classFocus === 'Good').length;
    const goodBehaviour = studentReports.filter(r => r.attendance && (r.classBehaviour === 'Good' || r.classBehaviour === 'Excellent')).length;

    return { presentDays, goodFocus, goodBehaviour };
  };

  // Per-student "More" menu state
  const [studentMoreMenuId, setStudentMoreMenuId] = useState<string | null>(null);

  // Form State - Report
  const [reportAttendance, setReportAttendance] = useState(true);
  const [reportHomework, setReportHomework] = useState<'Complete' | 'Incomplete' | 'Partially Complete'>('Complete');
  const [reportFocus, setReportFocus] = useState<'Good' | 'Moderate' | 'Poor' | 'Extremely Poor'>('Good');
  const [reportBehaviour, setReportBehaviour] = useState<'Excellent' | 'Good' | 'Normal' | 'Poor'>('Good');
  const [reportComplaint, setReportComplaint] = useState('');
  const [reportDate, setReportDate] = useState(new Date().toISOString().split('T')[0]);

  // Bulk Report State
  const [isBulkReportOpen, setIsBulkReportOpen] = useState(false);
  const [bulkReportStep, setBulkReportStep] = useState<'selectShift' | 'studentEntry'>('selectShift');
  const [bulkShift, setBulkShift] = useState<'1st' | '2nd' | '3rd' | null>(null);
  const [bulkStudentIndex, setBulkStudentIndex] = useState(0);
  const [bulkStudents, setBulkStudents] = useState<Student[]>([]);
  const [bulkReports, setBulkReports] = useState<DailyReport[]>([]);
  const [showShareSummary, setShowShareSummary] = useState(false);

  const reportTableRef = useRef<HTMLDivElement>(null);

  // Auth Listener
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // Firestore Students Listener
  useEffect(() => {
    if (!user) {
      setStudents([]);
      setStudentsLoading(false);
      return;
    }

    const studentsRef = collection(db, 'students');
    const q = query(
      studentsRef, 
      where('createdBy', '==', user.uid),
      orderBy('createdAt', 'desc')
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data(),
        createdAt: doc.data().createdAt instanceof Timestamp 
          ? doc.data().createdAt.toMillis() 
          : doc.data().createdAt
      })) as Student[];
      setStudents(data);
      setStudentsLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'students');
    });

    return () => unsubscribe();
  }, [user]);

  // Firestore Reports Listener
  useEffect(() => {
    if (!user) {
      setReports([]);
      setReportsLoading(false);
      return;
    }

    const reportsRef = collection(db, 'reports');
    const q = query(
      reportsRef, 
      where('createdBy', '==', user.uid),
      orderBy('createdAt', 'desc')
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data(),
        createdAt: doc.data().createdAt instanceof Timestamp 
          ? doc.data().createdAt.toMillis() 
          : doc.data().createdAt
      })) as DailyReport[];
      setReports(data);
      setReportsLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'reports');
    });

    return () => unsubscribe();
  }, [user]);

  const handleAddReport = async (e?: FormEvent) => {
    if (e) e.preventDefault();
    if (!reportingStudent || !user) return;

    try {
      const reportData: any = {
        studentId: reportingStudent.id,
        studentName: reportingStudent.name,
        date: reportDate,
        attendance: reportAttendance,
        createdBy: user.uid,
      };

      if (reportAttendance) {
        reportData.homeworkStatus = reportHomework;
        reportData.classFocus = reportFocus;
        reportData.classBehaviour = reportBehaviour;
        reportData.complaint = reportComplaint;
      } else {
        reportData.homeworkStatus = null;
        reportData.classFocus = null;
        reportData.classBehaviour = null;
        reportData.complaint = '';
      }

      let reportId = '';
      if (currentReport) {
        reportId = currentReport.id;
        await updateDoc(doc(db, 'reports', reportId), {
          ...reportData,
          updatedAt: serverTimestamp()
        });
      } else {
        const docRef = await addDoc(collection(db, 'reports'), {
          ...reportData,
          createdAt: serverTimestamp()
        });
        reportId = docRef.id;
      }

      const completedReport = { 
        id: reportId, 
        ...reportData, 
        createdAt: currentReport?.createdAt || Date.now() 
      } as DailyReport;

      // If it was a single report modal
      if (!isBulkReportOpen) {
        setReportingStudent(null);
        setCurrentReport(null);
        setIsReportModalOpen(false);
        resetReportForm();
      } else {
        setBulkReports(prev => {
          const exists = prev.findIndex(r => r.studentId === reportingStudent.id);
          if (exists !== -1) {
            const newArr = [...prev];
            newArr[exists] = completedReport;
            return newArr;
          }
          return [...prev, completedReport];
        });
        
        if (bulkStudentIndex < bulkStudents.length - 1) {
          const nextIndex = bulkStudentIndex + 1;
          const nextStudent = bulkStudents[nextIndex];
          setBulkStudentIndex(nextIndex);
          setReportingStudent(nextStudent);
          
          const existingReport = reports.find(r => r.studentId === nextStudent.id && r.date === reportDate);
          resetReportForm(existingReport);
        } else {
          setBulkReportStep('selectShift');
          setShowShareSummary(true);
        }
      }
    } catch (error) {
      handleFirestoreError(error, currentReport ? OperationType.UPDATE : OperationType.CREATE, 'reports');
    }
  };

  const shareMonthlyReport = async () => {
    if (!historyStudent) return;
    
    const doc = new jsPDF();
    const studentReports = reports.filter(r => 
      r.studentId === historyStudent.id && 
      r.date.startsWith(historyMonth)
    ).sort((a, b) => a.date.localeCompare(b.date));

    // Stats Calculation
    const totalSessions = studentReports.length;
    const presentDays = studentReports.filter(r => r.attendance).length;
    const goodFocusDays = studentReports.filter(r => r.attendance && r.classFocus === 'Good').length;
    const goodBehaviourDays = studentReports.filter(r => r.attendance && (r.classBehaviour === 'Good' || r.classBehaviour === 'Excellent')).length;

    // Overall Rating logic
    let rating = 'Moderate';
    if (presentDays > 0) {
      const focusScore = goodFocusDays / presentDays;
      const behaviorScore = goodBehaviourDays / presentDays;
      const combined = (focusScore + behaviorScore) / 2;
      
      if (combined > 0.85) rating = 'Excellent';
      else if (combined > 0.7) rating = 'Good';
      else if (combined > 0.4) rating = 'Moderate';
      else if (combined > 0.2) rating = 'Bad';
      else rating = 'Poor';
    } else {
      rating = 'No Data';
    }

    const monthLabel = new Date(historyMonth + '-01').toLocaleString('default', { month: 'long', year: 'numeric' });

    // PDF Header
    doc.setFontSize(22);
    doc.setTextColor(0, 122, 255);
    doc.text(`Monthly Progress Report`, 105, 15, { align: 'center' });
    
    doc.setFontSize(12);
    doc.setTextColor(100, 100, 100);
    doc.text(monthLabel, 200, 15, { align: 'right' });

    // Student Info
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text(`Student: ${historyStudent.name}`, 15, 30);
    doc.text(`Class: ${historyStudent.class}`, 15, 37);
    doc.text(`Contact: ${historyStudent.contactNumber}`, 15, 44);

    // Quick Stats Section
    doc.setFillColor(245, 247, 250);
    doc.rect(15, 52, 180, 42, 'F');
    
    doc.setFontSize(14);
    doc.setTextColor(0, 122, 255);
    doc.text("Quick Information Summary", 20, 60);
    
    doc.setFontSize(10);
    doc.setTextColor(50, 50, 50);
    doc.text(`Attendance: ${presentDays}/${totalSessions}`, 20, 72);
    doc.text(`Good behaviour: ${goodBehaviourDays} / out of ${presentDays} present days`, 20, 79);
    doc.text(`Focused in class: ${goodFocusDays} / out of ${presentDays} present days`, 20, 86);
    
    doc.text(`Overall all studies:`, 130, 72);
    doc.setFontSize(16);
    doc.setTextColor(0, 122, 255);
    doc.text(rating, 130, 84);

    const tableCols = ["Date", "Attendance", "Homework", "Focus", "Behaviour", "Complaints"];
    const tableRows = studentReports.map(r => [
      r.date,
      r.attendance ? "Present" : "Absent",
      r.attendance ? r.homeworkStatus || "-" : "-",
      r.attendance ? r.classFocus || "-" : "-",
      r.attendance ? r.classBehaviour || "-" : "-",
      r.attendance ? (r.complaint || "-") : "-"
    ]);

    autoTable(doc, {
      startY: 105,
      head: [tableCols],
      body: tableRows,
      theme: 'grid',
      headStyles: { fillColor: [0, 122, 255] },
      styles: { fontSize: 9 },
    });

    const pdfBlob = doc.output('blob');
    const file = new File([pdfBlob], `Report_${historyStudent.name}_${historyMonth}.pdf`, { type: 'application/pdf' });

    // WhatsApp logic
    const waNumber = historyStudent.contactNumber.replace(/\D/g, '');
    const cleanNumber = waNumber.startsWith('91') ? waNumber : `91${waNumber}`;
    
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: `Report for ${historyStudent.name}`,
          text: `Progress report for ${monthLabel} for ${historyStudent.name}`
        });
      } catch (error) {
        console.error('Sharing failed', error);
      }
    } else {
      doc.save(`Report_${historyStudent.name}_${historyMonth}.pdf`);
      // Open WhatsApp with text
      const text = encodeURIComponent(`Hello, here is the monthly report for ${historyStudent.name} for ${monthLabel}. I have downloaded the PDF and will send it shortly.`);
      window.open(`https://wa.me/${cleanNumber}?text=${text}`, '_blank');
    }
  };

  const getStatusColor = (type: 'hw' | 'focus' | 'behaviour' | 'attendance', value?: string | boolean) => {
    if (value === false || value === 'Absent') return 'bg-yellow-400 text-black';
    
    switch (value) {
      // HW
      case 'Complete': return 'bg-green-500 text-white';
      case 'Partially Complete': return 'bg-yellow-600 text-white';
      case 'Incomplete': return 'bg-red-500 text-white';
      // Focus
      case 'Good': return 'bg-green-500 text-white';
      case 'Moderate': return 'bg-yellow-600 text-white';
      case 'Poor':
      case 'Extremely Poor': return 'bg-red-500 text-white';
      // Behaviour
      case 'Excellent':
      case 'Good': return 'bg-green-500 text-white';
      case 'Normal': return 'bg-yellow-600 text-white';
      case 'Poor': return 'bg-red-500 text-white';
      default: return 'bg-gray-100 text-gray-500';
    }
  };

  const shareReportImage = async () => {
    if (!reportTableRef.current) return;
    
    try {
      const dataUrl = await toPng(reportTableRef.current, { 
        cacheBust: true, 
        backgroundColor: '#ffffff',
        style: {
          transform: 'scale(1)',
        }
      });
      
      const response = await fetch(dataUrl);
      const blob = await response.blob();
      const file = new File([blob], `Report_${bulkShift}_${reportDate}.png`, { type: 'image/png' });

      // Check if Web Share API is available and can share files
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({
          files: [file],
          title: `Daily Report - ${bulkShift} Shift`,
          text: `Today's report for ${bulkShift} Shift - ${reportDate}`
        });
      } else {
        // Fallback: Download and explain
        const link = document.createElement('a');
        link.download = `Report_${bulkShift}_${reportDate}.png`;
        link.href = dataUrl;
        link.click();
        alert('Sharing directly is not supported in this browser. The report image has been downloaded. You can now attach it to your WhatsApp group.');
      }
    } catch (error) {
      console.error('Error sharing image:', error);
      alert('Could not generate the report image. Please try again.');
    }
  };

  const resetReportForm = (existing?: DailyReport) => {
    if (existing) {
      setCurrentReport(existing);
      setReportAttendance(existing.attendance);
      setReportHomework(existing.homeworkStatus || 'Complete');
      setReportFocus(existing.classFocus || 'Good');
      setReportBehaviour(existing.classBehaviour || 'Good');
      setReportComplaint(existing.complaint || '');
    } else {
      setCurrentReport(null);
      setReportAttendance(true);
      setReportHomework('Complete');
      setReportFocus('Good');
      setReportBehaviour('Good');
      setReportComplaint('');
    }
  };

  const startBulkReport = (shift: '1st' | '2nd' | '3rd') => {
    const shiftStudents = students.filter(s => s.shift === shift);
    if (shiftStudents.length === 0) {
      alert(`No students found in ${shift} shift.`);
      return;
    }
    
    // Find already existing reports for this shift and today's date
    const today = new Date().toISOString().split('T')[0];
    setReportDate(today);
    
    const existingReportsInShift = reports.filter(r => 
      r.date === today && 
      shiftStudents.some(s => s.id === r.studentId)
    );

    setBulkShift(shift);
    setBulkStudents(shiftStudents);
    setBulkReports(existingReportsInShift); // Pre-populate already logged reports
    setBulkStudentIndex(0);
    const firstStudent = shiftStudents[0];
    setReportingStudent(firstStudent);
    setBulkReportStep('studentEntry');
    setIsBulkReportOpen(true);
    
    const firstExisting = existingReportsInShift.find(r => r.studentId === firstStudent.id);
    resetReportForm(firstExisting);
  };

  const exportToExcel = () => {
    if (reports.length === 0) return;
    
    const exportData = reports.map(r => ({
      Date: r.date,
      Student: r.studentName,
      Attendance: r.attendance ? 'Present' : 'Absent',
      'Homework Status': r.homeworkStatus || '-',
      'Class Focus': r.classFocus || '-',
      'Class Behaviour': r.classBehaviour || '-',
      Complaint: r.complaint || '-',
      'Logged At': new Date(r.createdAt).toLocaleString()
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Daily Reports");
    XLSX.writeFile(wb, `TutorReports_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const openReportModal = (student: Student) => {
    setReportingStudent(student);
    const existing = reports.find(r => r.studentId === student.id && r.date === reportDate);
    resetReportForm(existing);
    setIsReportModalOpen(true);
  };

  const handleClearTodayReports = async () => {
    try {
      const todayReports = reports.filter(r => r.date === reportDate);
      if (todayReports.length === 0) {
        alert('No reports found for ' + reportDate);
        setIsClearReportsConfirmOpen(false);
        return;
      }

      await Promise.all(todayReports.map(r => deleteDoc(doc(db, 'reports', r.id))));
      setIsClearReportsConfirmOpen(false);
      setIsReportsMoreMenuOpen(false);
      alert(`Cleared ${todayReports.length} reports for ${reportDate}`);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'reports');
    }
  };

  const handleAddOrUpdateStudent = async (e: FormEvent) => {
    e.preventDefault();
    if (!newName || !newContact || !user) return;

    try {
      if (editingStudent) {
        // Update existing
        await updateDoc(doc(db, 'students', editingStudent.id), {
          name: newName,
          class: newClass,
          shift: newShift,
          contactNumber: newContact,
          updatedAt: serverTimestamp(),
        });
      } else {
        // Create new
        await addDoc(collection(db, 'students'), {
          name: newName,
          class: newClass,
          shift: newShift,
          contactNumber: newContact,
          createdAt: serverTimestamp(),
          createdBy: user.uid,
        });
      }
      
      closeModal();
    } catch (error) {
      handleFirestoreError(error, editingStudent ? OperationType.UPDATE : OperationType.CREATE, 'students');
    }
  };

  const confirmDeleteStudent = (student: Student) => {
    setStudentToDelete(student);
    setIsDeleteConfirmOpen(true);
    setStudentMoreMenuId(null);
  };

  const handleDeleteStudent = async () => {
    if (!studentToDelete) return;
    try {
      await deleteDoc(doc(db, 'students', studentToDelete.id));
      setIsDeleteConfirmOpen(false);
      setStudentToDelete(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'students');
    }
  };

  const openEditModal = (student: Student) => {
    setEditingStudent(student);
    setNewName(student.name);
    setNewClass(student.class);
    setNewShift(student.shift || '1st');
    setNewContact(student.contactNumber);
    setIsAddModalOpen(true);
    setStudentMoreMenuId(null);
  };

  const closeModal = () => {
    setIsAddModalOpen(false);
    setEditingStudent(null);
    setNewName('');
    setNewClass('1st');
    setNewShift('1st');
    setNewContact('');
  };

  const filteredStudents = useMemo(() => {
    return students.filter(s => 
      s.name.toLowerCase().includes(searchQuery.toLowerCase()) || 
      s.class.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [students, searchQuery]);

  if (authLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <Loader2 className="w-8 h-8 text-[#007AFF] animate-spin" />
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-6">
        <div className="w-full max-w-sm bg-white rounded-3xl p-8 shadow-lg text-center space-y-8">
          <div className="space-y-4">
            <div className="w-20 h-20 bg-blue-50 rounded-3xl flex items-center justify-center mx-auto text-[#007AFF]">
              <GraduationCap className="w-10 h-10" />
            </div>
            <h1 className="text-3xl font-extrabold tracking-tight text-gray-900">ATC Reports</h1>
            <p className="text-gray-500 text-sm font-medium">Abhishek Tuition Classes<br/>Student Management Portal</p>
          </div>

          <button 
            onClick={() => loginWithGoogle()}
            className="w-full py-4 bg-[#007AFF] text-white rounded-2xl font-bold flex items-center justify-center gap-3 shadow-lg shadow-blue-500/20 active:scale-[0.98] transition-all"
          >
            <img src="https://www.google.com/favicon.ico" className="w-5 h-5 brightness-0 invert" alt="" />
            Sign in with Google
          </button>
          
          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-[0.2em]">Authorized Access Only</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-900 select-none flex flex-col lg:flex-row shadow-sm overflow-hidden">
      {/* Desktop Sidebar */}
      <aside className="hidden lg:flex flex-col w-72 bg-white border-r border-gray-100 h-screen sticky top-0 z-50">
        <div className="p-8">
          <div className="flex items-center gap-3 mb-10">
            <div className="w-10 h-10 bg-[#007AFF] rounded-xl flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
              <GraduationCap className="w-6 h-6" />
            </div>
            <h1 className="text-xl font-black text-gray-900 tracking-tight">ATC Reports</h1>
          </div>
          
          <nav className="space-y-2">
            {[
              { id: 'home' as TabType, label: 'Dashboard', icon: Home },
              { id: 'students' as TabType, label: 'Students', icon: Users },
              { id: 'reports' as TabType, label: 'Insights', icon: BarChart3 },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`w-full flex items-center gap-3 px-4 py-3.5 rounded-2xl font-bold text-sm transition-all ${
                  activeTab === tab.id 
                    ? 'bg-blue-50 text-[#007AFF] shadow-sm' 
                    : 'text-gray-400 hover:bg-gray-50 hover:text-gray-600'
                }`}
              >
                <tab.icon className="w-5 h-5" />
                {tab.label}
              </button>
            ))}
          </nav>
        </div>
        
        <div className="mt-auto p-8">
          <div className="flex items-center gap-3 mb-6 p-4 bg-gray-50 rounded-2xl overflow-hidden">
            {user.photoURL && <img src={user.photoURL} className="w-10 h-10 rounded-full" alt="" />}
            <div className="min-w-0">
              <p className="text-sm font-bold text-gray-900 truncate">{user.displayName || 'Authorized'}</p>
              <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest truncate">{user.email}</p>
            </div>
          </div>
          <button 
            onClick={() => logout()}
            className="w-full flex items-center gap-3 px-4 py-3.5 rounded-2xl font-bold text-sm text-red-500 hover:bg-red-50 transition-all"
          >
            <LogOut className="w-5 h-5" />
            Sign Out
          </button>
        </div>
      </aside>
      <div className="flex-1 flex flex-col min-w-0 h-screen overflow-y-auto">
        {/* Top Bar (Mobile & Desktop Header) */}
        <header className="sticky top-0 z-40 bg-white/80 backdrop-blur-md border-b border-gray-100 lg:static lg:bg-transparent lg:border-none lg:backdrop-blur-none">
          <div className="max-w-7xl mx-auto px-6 h-20 flex items-center justify-between">
            <div className="lg:hidden">
              <h1 className="text-2xl font-black text-gray-900 tracking-tight">ATC Reports</h1>
            </div>
            <div className="hidden lg:block">
              <h1 className="text-3xl font-black text-gray-900 tracking-tight capitalize">{activeTab} View</h1>
              <p className="text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] mt-1">Management Portal v2.0</p>
            </div>
            
            <div className="flex items-center gap-3">
              <button 
                onClick={() => setIsAddModalOpen(true)}
                className="p-2 sm:p-3 bg-[#007AFF] text-white rounded-2xl shadow-sm hover:shadow-md hover:-translate-y-0.5 active:translate-y-0 transition-all flex items-center gap-2"
              >
                <Plus className="w-5 h-5" />
                <span className="hidden sm:inline font-bold text-sm">New Student</span>
              </button>
            <div className="relative">
              <button 
                onClick={() => setIsMenuOpen(!isMenuOpen)}
                className="p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-400 overflow-hidden"
                id="more-menu-trigger"
              >
                {user.photoURL ? (
                  <img src={user.photoURL} className="w-5 h-5 rounded-full" alt="" />
                ) : (
                  <MoreVertical className="w-5 h-5" />
                )}
              </button>

              <AnimatePresence>
                {isMenuOpen && (
                  <>
                    <div 
                      className="fixed inset-0 z-40" 
                      onClick={() => setIsMenuOpen(false)} 
                    />
                    <motion.div 
                      initial={{ opacity: 0, scale: 0.95, y: -10 }}
                      animate={{ opacity: 1, scale: 1, y: 0 }}
                      exit={{ opacity: 0, scale: 0.95, y: -10 }}
                      className="absolute right-0 mt-2 w-56 bg-white rounded-2xl shadow-xl border border-gray-100 z-50 overflow-hidden"
                    >
                      <div className="p-4 border-b border-gray-50 flex items-center gap-3">
                        <div className="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center text-[#007AFF]">
                          {user.photoURL ? <img src={user.photoURL} className="rounded-full" alt="" /> : <UserIcon className="w-5 h-5" />}
                        </div>
                        <div className="overflow-hidden">
                          <p className="text-sm font-bold text-gray-900 truncate">{user.displayName || 'Abhishek'}</p>
                          <p className="text-[10px] text-gray-400 font-bold uppercase truncate">{user.email}</p>
                        </div>
                      </div>
                      <button 
                        onClick={() => {
                          setIsAddModalOpen(true);
                          setIsMenuOpen(false);
                        }}
                        className="w-full px-4 py-3 text-left flex items-center gap-3 hover:bg-gray-50 transition-colors"
                      >
                        <UserPlus className="w-4 h-4 text-[#007AFF]" />
                        <span className="text-sm font-semibold">New Registration</span>
                      </button>
                      <button 
                        onClick={() => {
                          logout();
                          setIsMenuOpen(false);
                        }}
                        className="w-full px-4 py-3 text-left flex items-center gap-3 hover:bg-red-50 transition-colors text-red-600"
                      >
                        <LogOut className="w-4 h-4" />
                        <span className="text-sm font-semibold">Sign Out</span>
                      </button>
                    </motion.div>
                  </>
                )}
              </AnimatePresence>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto w-full p-6 lg:p-10 space-y-8">
        <AnimatePresence mode="wait">
          {activeTab === 'home' && (
            <motion.div
              key="home"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6"
            >
              {/* Special Big Button for Daily Report */}
              <button 
                onClick={() => {
                  setBulkReportStep('selectShift');
                  setIsBulkReportOpen(true);
                }}
                className="col-span-full bg-gradient-to-br from-[#007AFF] to-[#0062CC] rounded-3xl p-8 sm:p-10 text-white flex flex-col items-center justify-center gap-6 shadow-md hover:shadow-lg transition-all group overflow-hidden relative min-h-[250px]"
              >
                <div className="absolute top-0 right-0 p-4 opacity-[0.05]">
                  <FileSpreadsheet className="w-48 h-48 sm:w-64 sm:h-64 -mr-10 -mt-10 sm:-mr-20 sm:-mt-20 group-hover:scale-105 transition-transform" />
                </div>
                <div className="w-20 h-20 bg-white/10 backdrop-blur-md rounded-2xl flex items-center justify-center group-hover:scale-105 transition-transform relative z-10">
                  <FileSpreadsheet className="w-10 h-10" />
                </div>
                <div className="text-center relative z-10">
                  <h3 className="text-2xl sm:text-3xl font-black tracking-tight">Add today's report</h3>
                  <p className="text-white/80 text-xs sm:text-sm font-bold uppercase tracking-widest mt-2">Daily Attendance & Progress Tracking</p>
                </div>
                <div className="px-6 py-2 sm:px-8 sm:py-3 bg-white text-[#007AFF] rounded-full font-black text-xs uppercase tracking-widest shadow-sm relative z-10">
                  Start Entry
                </div>
              </button>

              <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 flex flex-col justify-between group hover:shadow-md hover:-translate-y-0.5 transition-all">
                <div className="flex justify-between items-center mb-4">
                  <span className="text-gray-400 text-[10px] font-black uppercase tracking-widest">Enrolled Students</span>
                  <div className="p-3 bg-blue-50 text-[#007AFF] rounded-2xl">
                    <Users className="w-6 h-6" />
                  </div>
                </div>
                <div>
                  <p className="text-4xl font-black text-gray-900 leading-none">{studentsLoading ? '...' : students.length}</p>
                  <button 
                    onClick={() => setActiveTab('students')}
                    className="text-[#007AFF] text-[10px] font-black uppercase tracking-widest mt-3 hover:underline underline-offset-4"
                  >
                    Manage Directory &rarr;
                  </button>
                </div>
              </div>

              <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 flex flex-col justify-between group hover:shadow-md hover:-translate-y-0.5 transition-all">
                <div className="flex justify-between items-center mb-4">
                  <span className="text-gray-400 text-[10px] font-black uppercase tracking-widest">Reports Today</span>
                  <div className="p-3 bg-green-50 text-green-600 rounded-2xl">
                    <BarChart3 className="w-6 h-6" />
                  </div>
                </div>
                <div>
                  <p className="text-4xl font-black text-gray-900 leading-none">{reports.filter(r => r.date === reportDate).length}</p>
                  <button 
                    onClick={() => setActiveTab('reports')}
                    className="text-green-600 text-[10px] font-black uppercase tracking-widest mt-3 hover:underline underline-offset-4"
                  >
                    View Insights &rarr;
                  </button>
                </div>
              </div>

              <div className="space-y-2">
                <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest px-2">Quick Access</h3>
                <div className="bg-white rounded-2xl divide-y divide-gray-50 overflow-hidden shadow-sm border border-gray-50">
                  <button 
                    onClick={() => setIsAddModalOpen(true)}
                    className="w-full p-4 flex items-center justify-between hover:bg-gray-50 transition-colors text-left"
                  >
                    <div className="flex items-center gap-4">
                      <div className="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center text-[#007AFF]">
                        <UserPlus className="w-5 h-5" />
                      </div>
                      <p className="font-semibold text-gray-900">New Registration</p>
                    </div>
                    <ChevronRight className="w-4 h-4 text-gray-300" />
                  </button>
                  <button 
                    className="w-full p-4 flex items-center justify-between hover:bg-gray-50 transition-colors text-left"
                  >
                    <div className="flex items-center gap-4">
                      <div className="w-10 h-10 rounded-full bg-purple-50 flex items-center justify-center text-purple-600">
                        <GraduationCap className="w-5 h-5" />
                      </div>
                      <p className="font-semibold text-gray-900">Class Progress</p>
                    </div>
                    <ChevronRight className="w-4 h-4 text-gray-300" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'students' && (
            <motion.div
              key="students"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-4"
            >
              <div className="relative">
                <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                <input 
                  type="text"
                  placeholder="Search name or class..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-white border border-gray-100 rounded-2xl py-3.5 pl-11 pr-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#007AFF]/10 transition-all shadow-sm"
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-2 xl:grid-cols-3 gap-6">
                {studentsLoading ? (
                  <div className="p-12 text-center flex flex-col items-center gap-2">
                    <Loader2 className="w-6 h-6 text-[#007AFF] animate-spin" />
                    <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">Loading database...</p>
                  </div>
                ) : filteredStudents.length > 0 ? (
                  filteredStudents.map((student) => (
                    <motion.div
                      layout
                      key={student.id}
                      className="bg-white p-5 rounded-3xl border border-gray-100 shadow-sm hover:shadow hover:-translate-y-0.5 transition-all flex flex-col justify-between"
                    >
                      <div 
                        className="flex items-center gap-4 cursor-pointer"
                        onClick={() => setHistoryStudent(student)}
                      >
                        <div className="w-14 h-14 bg-gray-50 text-blue-600 rounded-2xl flex items-center justify-center font-black text-xl group-hover:bg-[#007AFF] group-hover:text-white transition-all shadow-inner">
                          {student.class.replace(/\D/g, '') || student.class.charAt(0)}
                        </div>
                        <div className="min-w-0">
                          <h3 className="font-black text-gray-900 text-lg truncate tracking-tight">{student.name}</h3>
                          <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mt-1">
                            Class {student.class} • {student.shift} Shift
                          </p>
                        </div>
                      </div>
                      
                      <div className="mt-6 pt-5 border-t border-gray-50 flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <a 
                            href={`tel:${student.contactNumber}`}
                            className="p-3 bg-gray-50 text-gray-400 rounded-2xl hover:bg-blue-50 hover:text-[#007AFF] transition-all"
                            onClick={(e) => e.stopPropagation()}
                          >
                            <Phone className="w-4 h-4" />
                          </a>
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              openReportModal(student);
                            }}
                            className="p-3 bg-blue-50 text-[#007AFF] rounded-2xl hover:bg-[#007AFF] hover:text-white transition-all"
                          >
                            <BarChart3 className="w-4 h-4" />
                          </button>
                        </div>
                        
                        <div className="relative">
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              setStudentMoreMenuId(studentMoreMenuId === student.id ? null : student.id);
                            }}
                            className="p-3 bg-gray-50 text-gray-400 rounded-2xl hover:bg-gray-100 transition-colors focus:outline-none"
                          >
                            <MoreVertical className="w-4 h-4" />
                          </button>
                          
                          <AnimatePresence>
                            {studentMoreMenuId === student.id && (
                              <>
                                <div className="fixed inset-0 z-40" onClick={() => setStudentMoreMenuId(null)} />
                                <motion.div
                                  initial={{ opacity: 0, scale: 0.95, y: -10 }}
                                  animate={{ opacity: 1, scale: 1, y: 0 }}
                                  exit={{ opacity: 0, scale: 0.95, y: -10 }}
                                  className="absolute right-0 bottom-full mb-2 w-36 bg-white rounded-2xl shadow-2xl border border-gray-100 z-50 overflow-hidden"
                                >
                                  <button 
                                    onClick={() => openEditModal(student)}
                                    className="w-full px-4 py-3 text-left text-xs font-bold text-gray-600 hover:bg-gray-50 flex items-center gap-2 uppercase tracking-widest"
                                  >
                                    <Edit2 className="w-3.5 h-3.5" />
                                    Edit
                                  </button>
                                  <button 
                                    onClick={() => confirmDeleteStudent(student)}
                                    className="w-full px-4 py-3 text-left text-xs font-bold text-red-600 hover:bg-red-50 flex items-center gap-2 uppercase tracking-widest"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                    Delete
                                  </button>
                                </motion.div>
                              </>
                            )}
                          </AnimatePresence>
                        </div>
                      </div>
                    </motion.div>
                  ))
                ) : (
                  <div className="p-12 text-center text-gray-400 italic text-sm">
                    {searchQuery ? 'No results found' : 'No students registered yet'}
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'reports' && (
            <motion.div
              key="reports"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <div className="bg-white rounded-3xl p-6 shadow-sm border border-gray-100">
                <div className="flex flex-col gap-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <h2 className="text-2xl font-black text-gray-900 tracking-tight">Reports & Insights</h2>
                      <p className="text-[10px] font-black text-gray-400 mt-0.5 uppercase tracking-widest">
                        {reportViewMode === 'day' ? `Daily Logs for ${reportDate}` : `Month Summary: ${new Date(reportDate).toLocaleString('default', { month: 'long', year: 'numeric' })}`}
                      </p>
                    </div>
                  </div>

                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    {/* View Switcher */}
                    <div className="p-1 bg-gray-100 rounded-2xl flex items-center self-start">
                      <button 
                        onClick={() => setReportViewMode('day')}
                        className={`px-4 py-2 rounded-xl text-[10px] font-bold uppercase tracking-wider transition-all ${reportViewMode === 'day' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-400'}`}
                      >
                        Day Wise
                      </button>
                      <button 
                        onClick={() => setReportViewMode('month')}
                        className={`px-4 py-2 rounded-xl text-[10px] font-bold uppercase tracking-wider transition-all ${reportViewMode === 'month' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-400'}`}
                      >
                        Month Wise
                      </button>
                    </div>

                    <div className="flex items-center gap-2">
                      <input 
                        type={reportViewMode === 'day' ? "date" : "month"} 
                        value={reportViewMode === 'day' ? reportDate : reportDate.slice(0, 7)}
                        onChange={(e) => setReportDate(reportViewMode === 'day' ? e.target.value : `${e.target.value}-01`)}
                        className="px-4 py-2.5 bg-gray-50 border-none rounded-xl font-bold text-gray-600 focus:ring-2 focus:ring-blue-500 transition-all outline-none text-xs"
                      />
                      <div className="flex items-center gap-1">
                        <button 
                          onClick={exportToExcel}
                          className="p-2.5 bg-green-50 text-green-600 rounded-xl hover:bg-green-100 transition-colors"
                          title="Export to Excel"
                        >
                          <FileSpreadsheet className="w-5 h-5" />
                        </button>
                        <div className="relative">
                          <button 
                            onClick={() => setIsReportsMoreMenuOpen(!isReportsMoreMenuOpen)}
                            className="p-2.5 bg-gray-50 text-gray-400 rounded-xl hover:bg-gray-100 transition-colors"
                          >
                            <MoreVertical className="w-5 h-5" />
                          </button>
                          
                          <AnimatePresence>
                            {isReportsMoreMenuOpen && (
                              <>
                                <div className="fixed inset-0 z-40" onClick={() => setIsReportsMoreMenuOpen(false)} />
                                <motion.div
                                  initial={{ opacity: 0, scale: 0.95, y: -10 }}
                                  animate={{ opacity: 1, scale: 1, y: 0 }}
                                  exit={{ opacity: 0, scale: 0.95, y: -10 }}
                                  className="absolute right-0 mt-2 w-48 bg-white rounded-xl shadow-xl border border-gray-100 z-50 overflow-hidden"
                                >
                                  <button 
                                    onClick={() => {
                                      setIsClearReportsConfirmOpen(true);
                                      setIsReportsMoreMenuOpen(false);
                                    }}
                                    className="w-full px-4 py-3 text-left text-sm font-semibold text-red-600 hover:bg-red-50 flex items-center gap-2"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                    Clear Day's Reports
                                  </button>
                                </motion.div>
                              </>
                            )}
                          </AnimatePresence>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {reportViewMode === 'day' ? (
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
                  {reportsLoading ? (
                    <div className="col-span-full p-12 text-center flex flex-col items-center gap-2 bg-white rounded-3xl border border-gray-50">
                      <Loader2 className="w-6 h-6 text-blue-600 animate-spin" />
                      <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">Fetching logs...</p>
                    </div>
                  ) : (
                    <>
                      {reports.filter(r => r.date === reportDate).length > 0 ? (
                        reports
                          .filter(r => r.date === reportDate)
                          .sort((a, b) => a.studentName.localeCompare(b.studentName))
                          .map((report) => (
                            <div key={report.id} className="bg-white p-6 rounded-3xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow">
                              <div className="flex justify-between items-start mb-4">
                                <div 
                                  className="cursor-pointer group"
                                  onClick={() => {
                                    const student = students.find(s => s.id === report.studentId);
                                    if (student) setHistoryStudent(student);
                                  }}
                                >
                                  <h3 className="font-bold text-gray-900 group-hover:text-blue-600 transition-colors">{report.studentName || 'Student'}</h3>
                                  <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">{report.date}</p>
                                </div>
                                <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${report.attendance ? 'bg-green-50 text-green-600' : 'bg-yellow-50 text-yellow-700'}`}>
                                  {report.attendance ? 'Present' : 'Absent'}
                                </span>
                              </div>
                              
                              {report.attendance && (
                                <div className="space-y-3">
                                  <div className="flex flex-wrap gap-2">
                                    <span className={`px-2.5 py-1 rounded-lg text-[10px] font-bold uppercase tracking-wider ${getStatusColor('hw', report.homeworkStatus)}`}>
                                      HW: {report.homeworkStatus}
                                    </span>
                                    <span className={`px-2.5 py-1 rounded-lg text-[10px] font-bold uppercase tracking-wider ${getStatusColor('focus', report.classFocus)}`}>
                                      Focus: {report.classFocus}
                                    </span>
                                    <span className={`px-2.5 py-1 rounded-lg text-[10px] font-bold uppercase tracking-wider ${getStatusColor('behaviour', report.classBehaviour)}`}>
                                      Behaviour: {report.classBehaviour}
                                    </span>
                                  </div>
                                  {report.complaint && (
                                    <div className="p-3 bg-red-50 rounded-xl border border-red-100/50">
                                      <p className="text-xs text-red-700 font-medium">
                                        <span className="font-black uppercase tracking-tighter mr-1 text-[10px]">Complaint:</span>
                                        {report.complaint}
                                      </p>
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          ))
                      ) : (
                        <div className="col-span-full p-20 text-center bg-white rounded-3xl border border-gray-50">
                          <div className="w-16 h-16 bg-gray-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-gray-200">
                            <BarChart3 className="w-8 h-8" />
                          </div>
                          <p className="text-gray-400 font-bold text-sm">No reports logged for this date</p>
                        </div>
                      )}
                    </>
                  )}
                </div>
              ) : (
                <div className="bg-white rounded-3xl overflow-hidden shadow-sm border border-gray-100">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-gray-50/50">
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Student</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-center border-b border-gray-50">Presence</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-center border-b border-gray-50">Good Focus</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-center border-b border-gray-50">Decent Behavior</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-50">
                        {students
                          .filter(s => isInCurrentShift(s))
                          .sort((a, b) => a.name.localeCompare(b.name))
                          .map(student => {
                            const stats = getMonthlyStats(student.id);
                            return (
                              <tr 
                                key={student.id} 
                                className="hover:bg-gray-50/50 transition-colors cursor-pointer"
                                onClick={() => setHistoryStudent(student)}
                              >
                                <td className="px-6 py-4">
                                  <div className="flex items-center gap-3">
                                    <div className="w-8 h-8 rounded-lg bg-blue-50 flex items-center justify-center text-blue-600 font-bold text-[10px]">
                                      {student.class.charAt(0)}
                                    </div>
                                    <span className="font-bold text-gray-900">{student.name}</span>
                                  </div>
                                </td>
                                <td className="px-6 py-4 text-center">
                                  <span className="px-3 py-1 bg-green-50 text-green-600 rounded-full text-[10px] font-black">
                                    {stats.presentDays} Days
                                  </span>
                                </td>
                                <td className="px-6 py-4 text-center">
                                  <span className="px-3 py-1 bg-blue-50 text-blue-600 rounded-full text-[10px] font-black">
                                    {stats.goodFocus} Days
                                  </span>
                                </td>
                                <td className="px-6 py-4 text-center">
                                  <span className="px-3 py-1 bg-purple-50 text-purple-600 rounded-full text-[10px] font-black">
                                    {stats.goodBehaviour} Days
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                        {students.filter(s => isInCurrentShift(s)).length === 0 && (
                          <tr>
                            <td colSpan={4} className="px-6 py-16 text-center text-gray-400 font-bold text-sm">
                              No students found in this shift
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Bottom Navigation */}
      <nav className="fixed bottom-0 left-0 right-0 z-40 bg-white/80 backdrop-blur-lg border-t border-gray-200">
        <div className="max-w-md mx-auto grid grid-cols-3 h-20 items-center px-4 pb-4">
          <button 
            onClick={() => setActiveTab('home')}
            className={`flex flex-col items-center gap-1 transition-all ${activeTab === 'home' ? 'text-[#007AFF]' : 'text-gray-400'}`}
          >
            <div className={`${activeTab === 'home' ? 'bg-blue-50 p-1.5 rounded-xl' : ''}`}>
               <Home className="w-6 h-6" fill={activeTab === 'home' ? 'currentColor' : 'none'} />
            </div>
            <span className="text-[10px] font-bold uppercase tracking-wider">Home</span>
          </button>
          <button 
            onClick={() => setActiveTab('students')}
            className={`flex flex-col items-center gap-1 transition-all ${activeTab === 'students' ? 'text-[#007AFF]' : 'text-gray-400'}`}
          >
            <div className={`${activeTab === 'students' ? 'bg-blue-50 p-1.5 rounded-xl' : ''}`}>
              <Users className="w-6 h-6" />
            </div>
            <span className="text-[10px] font-bold uppercase tracking-wider">Students</span>
          </button>
          <button 
            onClick={() => setActiveTab('reports')}
            className={`flex flex-col items-center gap-1 transition-all ${activeTab === 'reports' ? 'text-[#007AFF]' : 'text-gray-400'}`}
          >
            <div className={`${activeTab === 'reports' ? 'bg-blue-50 p-1.5 rounded-xl' : ''}`}>
              <BarChart3 className="w-6 h-6" />
            </div>
            <span className="text-[10px] font-bold uppercase tracking-wider">Reports</span>
          </button>
        </div>
        {/* iOS Home Indicator Placeholder */}
        <div className="h-1.5 w-32 bg-gray-900 rounded-full mx-auto mb-2 opacity-10" />
      </nav>

      {/* Add/Edit Student Modal - iOS Bottom Sheet Style */}
      <AnimatePresence>
        {isAddModalOpen && (
          <div className="fixed inset-0 z-50 flex items-end justify-center">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={closeModal}
              className="absolute inset-0 bg-black/30 backdrop-blur-[2px]"
            />
            <motion.div 
              initial={{ y: "100%" }}
              animate={{ y: 0 }}
              exit={{ y: "100%" }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="relative w-full max-w-md bg-white rounded-t-3xl p-6 shadow-2xl overflow-hidden pb-10"
            >
              <div className="w-12 h-1.5 bg-gray-200 rounded-full mx-auto mb-6" />
              
              <div className="flex items-center justify-between mb-8">
                <h2 className="text-xl font-bold text-gray-900">
                  {editingStudent ? 'Edit Student' : 'New Registration'}
                </h2>
                <button 
                  onClick={closeModal}
                  className="p-2 bg-gray-100 rounded-full text-gray-500 hover:bg-gray-200"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <form onSubmit={handleAddOrUpdateStudent} className="space-y-5">
                <div className="space-y-1.5">
                  <label className="block text-[11px] font-extrabold uppercase tracking-[0.1em] text-gray-400 ml-1">Student Name</label>
                  <input 
                    required
                    type="text"
                    placeholder="Full Name"
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    className="w-full bg-gray-50 border-none rounded-xl p-4 focus:ring-2 focus:ring-[#007AFF]/10 transition-all outline-none text-gray-900 font-medium"
                  />
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label className="block text-[11px] font-extrabold uppercase tracking-[0.1em] text-gray-400 ml-1">Class</label>
                    <div className="relative">
                      <select 
                        value={newClass}
                        onChange={(e) => setNewClass(e.target.value)}
                        className="w-full bg-gray-50 border-none rounded-xl p-4 appearance-none focus:ring-2 focus:ring-[#007AFF]/10 outline-none cursor-pointer text-gray-900 font-medium text-sm"
                      >
                        {CLASSES.map(c => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                      <ChevronRight className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 rotate-90 text-gray-400 pointer-events-none" />
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <label className="block text-[11px] font-extrabold uppercase tracking-[0.1em] text-gray-400 ml-1">Shift</label>
                    <div className="relative">
                      <select 
                        value={newShift}
                        onChange={(e) => setNewShift(e.target.value as any)}
                        className="w-full bg-gray-50 border-none rounded-xl p-4 appearance-none focus:ring-2 focus:ring-[#007AFF]/10 outline-none cursor-pointer text-gray-900 font-medium text-sm"
                      >
                        <option value="1st">1st Shift</option>
                        <option value="2nd">2nd Shift</option>
                        <option value="3rd">3rd Shift</option>
                      </select>
                      <ChevronRight className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 rotate-90 text-gray-400 pointer-events-none" />
                    </div>
                  </div>
                </div>

                <div className="space-y-1.5">
                  <label className="block text-[11px] font-extrabold uppercase tracking-[0.1em] text-gray-400 ml-1">Contact Number</label>
                  <input 
                    required
                    type="tel"
                    placeholder="+91"
                    value={newContact}
                    onChange={(e) => setNewContact(e.target.value)}
                    className="w-full bg-gray-50 border-none rounded-xl p-4 focus:ring-2 focus:ring-[#007AFF]/10 transition-all outline-none text-gray-900 font-medium"
                  />
                </div>

                <button 
                  type="submit"
                  className="w-full py-4.5 bg-[#007AFF] text-white rounded-2xl font-bold flex items-center justify-center gap-2 shadow-lg shadow-blue-500/10 active:scale-[0.98] transition-all mt-4"
                >
                  {editingStudent ? <Edit2 className="w-5 h-5" /> : <Plus className="w-5 h-5" />}
                  {editingStudent ? 'Update Details' : 'Register Student'}
                </button>
              </form>
            </motion.div>
          </div>
        )}

        {/* Bulk Daily Report Modal */}
        {isBulkReportOpen && (
          <div className="fixed inset-0 z-[70] flex items-end justify-center">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsBulkReportOpen(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-md"
            />
            <motion.div 
              initial={{ y: "100%" }}
              animate={{ y: 0 }}
              exit={{ y: "100%" }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="relative w-full max-w-lg bg-white rounded-t-3xl p-8 shadow-2xl overflow-hidden pb-12 overflow-y-auto max-h-[90vh]"
            >
              <div className="w-12 h-1.5 bg-gray-200 rounded-full mx-auto mb-8" />
              
              {bulkReportStep === 'selectShift' ? (
                <div className="space-y-8">
                  <div className="text-center">
                    <h2 className="text-2xl font-bold text-gray-900">Select Shift</h2>
                    <p className="text-sm text-gray-500 font-medium">Which shift are you reporting for today?</p>
                  </div>
                  
                  <div className="grid grid-cols-1 gap-4">
                    {['1st', '2nd', '3rd'].map((shift) => (
                      <button
                        key={shift}
                        onClick={() => startBulkReport(shift as any)}
                        className="w-full p-6 bg-gray-50 rounded-3xl flex items-center justify-between hover:bg-blue-50 transition-colors group"
                      >
                        <div className="flex items-center gap-4">
                          <div className="w-12 h-12 bg-white rounded-2xl flex items-center justify-center text-[#007AFF] font-bold group-hover:bg-[#007AFF] group-hover:text-white transition-all">
                            {shift[0]}
                          </div>
                          <div className="text-left">
                            <p className="font-bold text-gray-900">{shift} Shift</p>
                            <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">
                              {students.filter(s => s.shift === shift).length} Students
                            </p>
                          </div>
                        </div>
                        <ChevronRight className="w-5 h-5 text-gray-300 group-hover:text-[#007AFF]" />
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="space-y-8">
                  <div className="flex items-center justify-between border-b border-gray-50 pb-6 mb-2">
                    <div className="flex items-center gap-4">
                      <button 
                        onClick={() => setBulkReportStep('selectShift')}
                        className="p-2 bg-gray-100 rounded-full text-gray-500 hover:bg-gray-200"
                      >
                        <ArrowLeft className="w-4 h-4" />
                      </button>
                      <div>
                        <h2 className="text-xl font-bold text-gray-900">Student {bulkStudentIndex + 1}/{bulkStudents.length}</h2>
                        <p className="text-xs text-[#007AFF] font-bold uppercase tracking-widest">{bulkShift} Shift</p>
                      </div>
                    </div>
                    <div className="text-right">
                      <p className="text-sm font-bold text-gray-900">{reportingStudent?.name}</p>
                      <p className="text-[10px] text-gray-400 font-bold uppercase tracking-wider">{reportDate}</p>
                    </div>
                  </div>

                  <form onSubmit={handleAddReport} className="space-y-6">
                    <div className="space-y-2">
                      <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Attendance</label>
                      <div className="flex bg-gray-100 p-1.5 rounded-2xl gap-1.5">
                        <button
                          type="button"
                          onClick={() => setReportAttendance(true)}
                          className={`flex-1 py-3 px-4 rounded-xl font-bold text-sm flex items-center justify-center gap-2 transition-all ${reportAttendance ? 'bg-white text-[#007AFF] shadow-sm' : 'text-gray-500 hover:bg-gray-200'}`}
                        >
                          <CheckCircle2 className={`w-4 h-4 ${reportAttendance ? 'text-[#007AFF]' : 'text-gray-300'}`} />
                          Present
                        </button>
                        <button
                          type="button"
                          onClick={() => setReportAttendance(false)}
                          className={`flex-1 py-3 px-4 rounded-xl font-bold text-sm flex items-center justify-center gap-2 transition-all ${!reportAttendance ? 'bg-white text-red-500 shadow-sm' : 'text-gray-500 hover:bg-gray-200'}`}
                        >
                          <AlertCircle className={`w-4 h-4 ${!reportAttendance ? 'text-red-500' : 'text-gray-300'}`} />
                          Absent
                        </button>
                      </div>
                    </div>

                    <AnimatePresence mode="wait">
                      {reportAttendance && (
                        <motion.div
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: 'auto' }}
                          exit={{ opacity: 0, height: 0 }}
                          className="space-y-6 overflow-hidden"
                        >
                          <div className="space-y-2">
                            <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Homework Status</label>
                            <div className="flex flex-wrap gap-2">
                              {['Complete', 'Incomplete', 'Partially Complete'].map(status => (
                                <button
                                  key={status}
                                  type="button"
                                  onClick={() => setReportHomework(status as any)}
                                  className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${reportHomework === status ? getStatusColor('hw', status) + ' shadow-md' : 'bg-gray-100 text-gray-500 border border-transparent'}`}
                                >
                                  {status}
                                </button>
                              ))}
                            </div>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Class Focus</label>
                            <div className="flex flex-wrap gap-2">
                              {['Good', 'Moderate', 'Poor', 'Extremely Poor'].map(option => (
                                <button
                                  key={option}
                                  type="button"
                                  onClick={() => setReportFocus(option as any)}
                                  className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${reportFocus === option ? getStatusColor('focus', option) + ' shadow-md' : 'bg-gray-100 text-gray-500 border border-transparent'}`}
                                >
                                  {option}
                                </button>
                              ))}
                            </div>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Class Behaviour</label>
                            <div className="flex flex-wrap gap-2">
                              {['Excellent', 'Good', 'Normal', 'Poor'].map(option => (
                                <button
                                  key={option}
                                  type="button"
                                  onClick={() => setReportBehaviour(option as any)}
                                  className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${reportBehaviour === option ? getStatusColor('behaviour', option) + ' shadow-md' : 'bg-gray-100 text-gray-500 border border-transparent'}`}
                                >
                                  {option}
                                </button>
                              ))}
                            </div>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Any Complaint (Optional)</label>
                            <textarea
                              value={reportComplaint}
                              onChange={(e) => setReportComplaint(e.target.value)}
                              placeholder="e.g. Needs more focus on vocabulary"
                              className="w-full bg-gray-50 border border-gray-100 rounded-2xl p-4 text-sm font-medium outline-none focus:ring-2 focus:ring-blue-100 resize-none h-24"
                            />
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>

                    <button 
                      type="submit"
                      className="w-full py-5 bg-[#007AFF] text-white rounded-2xl font-bold text-lg shadow-md shadow-blue-500/20 flex items-center justify-center gap-3 active:scale-95 transition-all"
                    >
                      {bulkStudentIndex === bulkStudents.length - 1 
                        ? (currentReport ? 'Update & Finish' : 'Finish & Save Reports') 
                        : (currentReport ? 'Update & Next Student' : 'Save & Next Student')}
                      <ChevronRight className="w-5 h-5" />
                    </button>
                  </form>
                </div>
              )}
            </motion.div>
          </div>
        )}

        {/* Existing Daily Report Modal (Individual) */}
        {isReportModalOpen && reportingStudent && (
          <div className="fixed inset-0 z-50 flex items-end justify-center">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsReportModalOpen(false)}
              className="absolute inset-0 bg-black/30 backdrop-blur-[2px]"
            />
            <motion.div 
              initial={{ y: "100%" }}
              animate={{ y: 0 }}
              exit={{ y: "100%" }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="relative w-full max-w-md bg-white rounded-t-3xl p-6 shadow-2xl overflow-hidden pb-10 max-h-[90vh] overflow-y-auto"
            >
              <div className="w-12 h-1.5 bg-gray-200 rounded-full mx-auto mb-6" />
              
              <div className="flex items-center justify-between mb-8">
                <div>
                  <h2 className="text-xl font-bold text-gray-900">Individual Report</h2>
                  <p className="text-xs text-gray-500 font-bold uppercase tracking-widest">{reportingStudent.name}</p>
                </div>
                <button 
                  onClick={() => setIsReportModalOpen(false)}
                  className="p-2 bg-gray-100 rounded-full text-gray-500 hover:bg-gray-200"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <form onSubmit={handleAddReport} className="space-y-6">
                <div className="space-y-1.5">
                  <label className="block text-[11px] font-extrabold uppercase tracking-[0.1em] text-gray-400 ml-1">Report Date</label>
                  <input 
                    required
                    type="date"
                    value={reportDate}
                    onChange={(e) => setReportDate(e.target.value)}
                    className="w-full bg-gray-50 border-none rounded-xl p-4 outline-none text-gray-900 font-medium"
                  />
                </div>

                <div className="space-y-2">
                  <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Attendance</label>
                  <div className="flex bg-gray-100 p-1.5 rounded-2xl gap-1.5">
                    <button
                      type="button"
                      onClick={() => setReportAttendance(true)}
                      className={`flex-1 py-3 px-4 rounded-xl font-bold text-sm flex items-center justify-center gap-2 transition-all ${reportAttendance ? 'bg-white text-[#007AFF] shadow-sm' : 'text-gray-500 hover:bg-gray-200'}`}
                    >
                      Present
                    </button>
                    <button
                      type="button"
                      onClick={() => setReportAttendance(false)}
                      className={`flex-1 py-3 px-4 rounded-xl font-bold text-sm flex items-center justify-center gap-2 transition-all ${!reportAttendance ? 'bg-white text-red-500 shadow-sm' : 'text-gray-500 hover:bg-gray-200'}`}
                    >
                      Absent
                    </button>
                  </div>
                </div>

                <AnimatePresence mode="wait">
                  {reportAttendance && (
                    <motion.div
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="space-y-6 overflow-hidden"
                    >
                      <div className="space-y-2">
                        <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Homework Status</label>
                        <div className="flex flex-wrap gap-2">
                          {['Complete', 'Incomplete', 'Partially Complete'].map(status => (
                            <button
                              key={status}
                              type="button"
                              onClick={() => setReportHomework(status as any)}
                              className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${reportHomework === status ? getStatusColor('hw', status) + ' shadow-md' : 'bg-gray-100 text-gray-500 border border-transparent'}`}
                            >
                              {status}
                            </button>
                          ))}
                        </div>
                      </div>

                      <div className="space-y-2">
                        <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Class Focus</label>
                        <div className="flex flex-wrap gap-2">
                          {['Good', 'Moderate', 'Poor', 'Extremely Poor'].map(option => (
                            <button
                              key={option}
                              type="button"
                              onClick={() => setReportFocus(option as any)}
                              className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${reportFocus === option ? getStatusColor('focus', option) + ' shadow-md' : 'bg-gray-100 text-gray-500 border border-transparent'}`}
                            >
                              {option}
                            </button>
                          ))}
                        </div>
                      </div>

                      <div className="space-y-2">
                        <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Class Behaviour</label>
                        <div className="flex flex-wrap gap-2">
                          {['Excellent', 'Good', 'Normal', 'Poor'].map(option => (
                            <button
                              key={option}
                              type="button"
                              onClick={() => setReportBehaviour(option as any)}
                              className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all ${reportBehaviour === option ? getStatusColor('behaviour', option) + ' shadow-md' : 'bg-gray-100 text-gray-500 border border-transparent'}`}
                            >
                              {option}
                            </button>
                          ))}
                        </div>
                      </div>

                      <div className="space-y-2">
                        <label className="block text-[11px] font-extrabold uppercase tracking-[0.15em] text-gray-400 ml-1">Any Complaint (Optional)</label>
                        <textarea
                          value={reportComplaint}
                          onChange={(e) => setReportComplaint(e.target.value)}
                          placeholder="e.g. Needs more focus on vocabulary"
                          className="w-full bg-gray-50 border border-gray-100 rounded-2xl p-4 text-sm font-medium outline-none focus:ring-2 focus:ring-blue-100 resize-none h-24"
                        />
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>

                <button 
                  type="submit"
                  className="w-full py-4.5 bg-[#007AFF] text-white rounded-2xl font-bold flex items-center justify-center gap-2 shadow-lg shadow-blue-500/10 active:scale-[0.98] transition-all mt-4"
                >
                  {currentReport ? 'Update Report' : 'Save Report'}
                </button>
              </form>
            </motion.div>
          </div>
        )}
        {/* Delete Confirmation Modal */}
        {isDeleteConfirmOpen && studentToDelete && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsDeleteConfirmOpen(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-sm bg-white rounded-3xl p-8 shadow-xl text-center"
            >
              <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-6">
                <Trash2 className="w-8 h-8" />
              </div>
              <h2 className="text-xl font-bold text-gray-900 mb-2">Delete Student?</h2>
              <p className="text-sm text-gray-500 mb-8 font-medium">
                Are you sure you want to remove <span className="text-gray-900 font-bold">{studentToDelete.name}</span>? This action cannot be undone.
              </p>
              <div className="flex flex-col gap-3">
                <button 
                  onClick={handleDeleteStudent}
                  className="w-full py-4 bg-red-500 text-white rounded-2xl font-bold active:scale-[0.98] transition-all"
                >
                  Confirm Delete
                </button>
                <button 
                  onClick={() => setIsDeleteConfirmOpen(false)}
                  className="w-full py-4 bg-gray-100 text-gray-600 rounded-2xl font-bold active:scale-[0.98] transition-all"
                >
                  Cancel
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Student History Modal */}
        {historyStudent && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center sm:p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setHistoryStudent(null)}
              className="absolute inset-0 bg-black/60 backdrop-blur-md"
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="relative w-full max-w-4xl max-h-[90vh] bg-gray-50 rounded-3xl shadow-xl flex flex-col overflow-hidden"
            >
              {/* Header */}
              <div className="p-8 bg-white border-b border-gray-100 flex flex-col sm:flex-row sm:items-center justify-between gap-6 shrink-0">
                <div className="flex items-center gap-5">
                  <div className="w-20 h-20 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center font-black text-3xl shadow-sm border border-blue-100">
                    {historyStudent.name.charAt(0)}
                  </div>
                  <div>
                    <h2 className="text-3xl font-black text-gray-900 tracking-tight">{historyStudent.name}</h2>
                    <div className="flex flex-wrap gap-x-4 gap-y-1 mt-1 text-sm font-bold text-gray-400 uppercase tracking-widest">
                      <span>Class {historyStudent.class}</span>
                      <span>•</span>
                      <span className="text-blue-600">{historyStudent.contactNumber}</span>
                    </div>
                  </div>
                </div>
                
                <div className="flex items-center gap-3">
                  <input 
                    type="month" 
                    value={historyMonth}
                    onChange={(e) => setHistoryMonth(e.target.value)}
                    className="px-4 py-3 bg-gray-50 border-none rounded-2xl font-bold text-gray-600 focus:ring-2 focus:ring-blue-500 transition-all outline-none"
                  />
                  <button 
                    onClick={() => setHistoryStudent(null)}
                    className="p-4 bg-gray-100 text-gray-400 rounded-2xl hover:bg-gray-200 transition-colors"
                  >
                    <X className="w-6 h-6" />
                  </button>
                </div>
              </div>

              {/* Monthly Table */}
              <div className="flex-1 overflow-auto p-6">
                <div className="bg-white rounded-3xl overflow-hidden shadow-sm border border-gray-100">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-gray-50/50">
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Date</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Attendance</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Homework</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Focus</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Behaviour</th>
                          <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest border-b border-gray-50">Complaints</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-50">
                        {reports
                          .filter(r => r.studentId === historyStudent.id && r.date.startsWith(historyMonth))
                          .sort((a, b) => b.date.localeCompare(a.date))
                          .map((report) => (
                            <tr key={report.id} className="hover:bg-gray-50/50 transition-colors">
                              <td className="px-6 py-4 whitespace-nowrap">
                                <span className="font-bold text-gray-900">{new Date(report.date).toLocaleDateString(undefined, { day: 'numeric', month: 'short' })}</span>
                              </td>
                              <td className="px-6 py-4">
                                <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${report.attendance ? 'bg-green-50 text-green-600' : 'bg-red-50 text-red-600'}`}>
                                  {report.attendance ? 'Present' : 'Absent'}
                                </span>
                              </td>
                              <td className="px-6 py-4">
                                {report.attendance ? (
                                  <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${getStatusColor('hw', report.homeworkStatus)}`}>
                                    {report.homeworkStatus}
                                  </span>
                                ) : '-'}
                              </td>
                              <td className="px-6 py-4">
                                {report.attendance ? (
                                  <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${getStatusColor('focus', report.classFocus)}`}>
                                    {report.classFocus}
                                  </span>
                                ) : '-'}
                              </td>
                              <td className="px-6 py-4">
                                {report.attendance ? (
                                  <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${getStatusColor('behaviour', report.classBehaviour)}`}>
                                    {report.classBehaviour}
                                  </span>
                                ) : '-'}
                              </td>
                              <td className="px-6 py-4 max-w-[200px]">
                                <p className="text-sm text-gray-500 font-medium truncate" title={report.complaint}>
                                  {report.complaint || '-'}
                                </p>
                              </td>
                            </tr>
                          ))}
                        {reports.filter(r => r.studentId === historyStudent.id && r.date.startsWith(historyMonth)).length === 0 && (
                          <tr>
                            <td colSpan={6} className="px-6 py-20 text-center">
                              <div className="w-20 h-20 bg-gray-50 text-gray-200 rounded-3xl flex items-center justify-center mx-auto mb-4">
                                <FileText className="w-10 h-10" />
                              </div>
                              <p className="text-gray-400 font-bold">No reports found for this month</p>
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>

              {/* Action Bar */}
              <div className="p-8 bg-white border-t border-gray-100 flex items-center justify-center shrink-0">
                <button 
                  onClick={shareMonthlyReport}
                  className="w-full max-w-sm py-5 bg-[#25D366] text-white rounded-2xl font-black text-lg shadow-md shadow-green-500/20 flex items-center justify-center gap-3 active:scale-95 transition-all"
                >
                  <FileText className="w-6 h-6" />
                  Share Monthly Report (PDF)
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Clear Reports Confirmation Modal */}
        {isClearReportsConfirmOpen && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsClearReportsConfirmOpen(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-sm bg-white rounded-3xl p-8 shadow-xl text-center"
            >
              <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-6">
                <AlertCircle className="w-8 h-8" />
              </div>
              <h2 className="text-xl font-bold text-gray-900 mb-2">Clear All Reports?</h2>
              <p className="text-sm text-gray-500 mb-8 font-medium">
                This will delete all student reports for <span className="text-gray-900 font-bold">{reportDate}</span>. This action cannot be undone.
              </p>
              <div className="flex flex-col gap-3">
                <button 
                  onClick={handleClearTodayReports}
                  className="w-full py-4 bg-red-500 text-white rounded-2xl font-bold active:scale-[0.98] transition-all"
                >
                  Delete All for Today
                </button>
                <button 
                  onClick={() => setIsClearReportsConfirmOpen(false)}
                  className="w-full py-4 bg-gray-100 text-gray-600 rounded-2xl font-bold active:scale-[0.98] transition-all"
                >
                  Cancel
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
      <div className="fixed -left-[2000px] top-0 pointer-events-none">
        <div 
          ref={reportTableRef} 
          className="p-8 bg-white w-[1000px]"
        >
          <div className="text-center mb-8 border-b-4 border-blue-600 pb-6">
            <h1 className="text-4xl font-black text-gray-900 uppercase tracking-widest">Daily Report - {bulkShift} Shift</h1>
            <p className="text-2xl font-bold text-blue-600 mt-2">{reportDate}</p>
          </div>
          
          <div className="overflow-hidden rounded-2xl border-2 border-gray-200">
            <table className="w-full border-collapse">
              <thead>
                <tr>
                  <th className="border-2 border-gray-200 p-4 bg-gray-50 text-gray-800 font-black uppercase text-sm text-left">Student Name</th>
                  <th className="border-2 border-gray-200 p-4 bg-gray-50 text-gray-800 font-black uppercase text-sm text-center">Homework</th>
                  <th className="border-2 border-gray-200 p-4 bg-gray-50 text-gray-800 font-black uppercase text-sm text-center">Focus</th>
                  <th className="border-2 border-gray-200 p-4 bg-gray-50 text-gray-800 font-black uppercase text-sm text-center">Behaviour</th>
                  <th className="border-2 border-gray-200 p-4 bg-gray-50 text-gray-800 font-black uppercase text-sm text-center">Complaints</th>
                </tr>
              </thead>
              <tbody>
                {bulkReports.map(r => (
                  <tr key={r.id}>
                    <td className="border-2 border-gray-200 p-4 font-black text-gray-900">{r.studentName}</td>
                    <td className={`border-2 border-gray-200 p-4 text-center font-bold text-sm ${getStatusColor('hw', r.attendance ? r.homeworkStatus : 'Absent')}`}>
                      {r.attendance ? r.homeworkStatus : 'ABSENT'}
                    </td>
                    <td className={`border-2 border-gray-200 p-4 text-center font-bold text-sm ${getStatusColor('focus', r.attendance ? r.classFocus : 'Absent')}`}>
                      {r.attendance ? r.classFocus : 'ABSENT'}
                    </td>
                    <td className={`border-2 border-gray-200 p-4 text-center font-bold text-sm ${getStatusColor('behaviour', r.attendance ? r.classBehaviour : 'Absent')}`}>
                      {r.attendance ? r.classBehaviour : 'ABSENT'}
                    </td>
                    <td className="border-2 border-gray-200 p-4 text-center font-bold text-xs text-gray-600 max-w-[200px] break-words">
                      {r.attendance ? (r.complaint || '-') : 'ABSENT'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          
          <div className="mt-8 pt-6 border-t-2 border-gray-100 flex justify-end items-center text-gray-400 font-black uppercase tracking-widest text-[10px]">
            <p>{new Date().toLocaleTimeString()}</p>
          </div>
        </div>
      </div>

      {/* Share Summary Modal */}
      <AnimatePresence>
        {showShareSummary && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-black/60 backdrop-blur-md"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative w-full max-w-sm bg-white rounded-3xl p-8 shadow-xl text-center"
            >
              <div className="w-20 h-20 bg-green-50 text-green-500 rounded-3xl flex items-center justify-center mx-auto mb-6 shadow-sm ring-1 ring-green-100">
                <CheckCircle2 className="w-10 h-10" />
              </div>
              
              <h2 className="text-2xl font-bold text-gray-900 mb-2">Reports Saved!</h2>
              <p className="text-sm text-gray-500 mb-8 font-medium">
                All {bulkStudents.length} reports for <span className="font-bold text-[#007AFF] uppercase">{bulkShift} Shift</span> have been successfully logged.
              </p>
              
              <div className="flex flex-col gap-3">
                <button 
                  onClick={shareReportImage}
                  className="w-full py-5 bg-[#25D366] text-white rounded-2xl font-bold flex items-center justify-center gap-3 shadow-xl shadow-green-500/20 active:scale-[0.98] transition-all"
                >
                  <FileSpreadsheet className="w-5 h-5" />
                  Share to WhatsApp
                </button>
                <button 
                  onClick={() => {
                    setShowShareSummary(false);
                    setIsBulkReportOpen(false);
                    setBulkShift(null);
                    setBulkStudents([]);
                    setBulkReports([]);
                  }}
                  className="w-full py-4 bg-gray-100 text-gray-600 rounded-2xl font-bold active:scale-[0.98] transition-all"
                >
                  Close & Go Home
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
        {/* Mobile Navigation Bar */}
        <nav className="fixed bottom-0 left-0 right-0 z-40 p-4 pb-8 lg:hidden">
          <div className="max-w-md mx-auto bg-white/80 backdrop-blur-md border border-gray-100 rounded-3xl p-2 flex items-center justify-between shadow-lg shadow-blue-500/10">
            {[
              { id: 'home' as TabType, label: 'Home', icon: Home },
              { id: 'students' as TabType, label: 'Students', icon: Users },
              { id: 'reports' as TabType, label: 'Reports', icon: BarChart3 },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`relative flex flex-col items-center justify-center w-20 h-14 rounded-2xl transition-all ${
                  activeTab === tab.id ? 'text-[#007AFF]' : 'text-gray-400'
                }`}
              >
                {activeTab === tab.id && (
                  <motion.div 
                    layoutId="activeTabMobile"
                    className="absolute inset-0 bg-blue-50 rounded-2xl -z-10"
                  />
                )}
                <tab.icon className="w-5 h-5 mb-1" />
                <span className="text-[10px] font-black uppercase tracking-widest">{tab.label}</span>
              </button>
            ))}
          </div>
        </nav>
      </div>
    </div>
  );
}
