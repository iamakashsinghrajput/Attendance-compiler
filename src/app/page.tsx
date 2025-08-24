// app/page.tsx
'use client';

import { useState, useEffect } from 'react';
import { Mail, Upload, Book, Building, Moon, ChevronDown, FileText, Copy, Check, Users, AlertCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

type FileState = File | null;

interface StudentData {
  rank: number;
  name: string;
  rollNumber: string;
  percentage: number;
}

interface EmailData {
  displayVersion: string;
  htmlVersion: string;
  plainTextVersion: string;
}

interface AttendanceDate {
  date: string;
  fullText: string;
  columnIndex: number;
}

interface StudentAttendance {
  name: string;
  email: string;
  status: 'present' | 'absent';
}

interface InternReportItem {
  id: number;
  content: string;
  type: string;
}

interface AttendanceStats {
  date: string;
  totalStudents: number;
  present: number;
  absent: number;
  presentPercentage: number;
  absentPercentage: number;
  presentStudents: StudentAttendance[];
  absentStudents: StudentAttendance[];
}

interface BatchData {
  id: number;
  name: string;
  total: number;
  present: number;
  absent: number;
}

interface EmailTemplate {
  trainingDate: string;
  batches: BatchData[];
  sheetsLink: string;
  generatedContent: string;
  plainTextContent?: string;
}

const colleges = [
  { id: 'srm', name: 'SRM' },
  { id: 'karpagam', name: 'Karpagam' },
  { id: 'niet', name: 'NIET' }
];

export default function EmailGeneratorPage() {
  const [selectedCollege, setSelectedCollege] = useState<string>('');
  const [attendanceFile, setAttendanceFile] = useState<FileState>(null);
  const [studentData, setStudentData] = useState<StudentData[]>([]);
  const [selectedStudents, setSelectedStudents] = useState<Set<number>>(new Set());
  const [generatedEmails, setGeneratedEmails] = useState<EmailData[]>([]);
  const [copiedEmailIndex, setCopiedEmailIndex] = useState<number | null>(null);
  const [copiedSubject, setCopiedSubject] = useState<boolean>(false);
  const [emailsSent, setEmailsSent] = useState<Set<number>>(new Set());
  // NIET attendance analysis states
  const [attendanceDates, setAttendanceDates] = useState<AttendanceDate[]>([]);
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [attendanceStats, setAttendanceStats] = useState<AttendanceStats | null>(null);
  const [rawAttendanceData, setRawAttendanceData] = useState<(string | number)[][]>([]);
  const [copiedPresentEmails, setCopiedPresentEmails] = useState<boolean>(false);
  const [copiedAbsentEmails, setCopiedAbsentEmails] = useState<boolean>(false);
  // Intern report states
  const [internReport, setInternReport] = useState<string>('');
  const [processedInternData, setProcessedInternData] = useState<InternReportItem[]>([]);
  const [internReportExpanded, setInternReportExpanded] = useState<boolean>(false);
  // Excel sheet selection states
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [selectedSheetsForProcessing, setSelectedSheetsForProcessing] = useState<Set<string>>(new Set());
  const [allSheetsAttendanceData, setAllSheetsAttendanceData] = useState<Map<string, AttendanceStats>>(new Map());
  // Email template states
  const [emailTemplate, setEmailTemplate] = useState<EmailTemplate>({
    trainingDate: '',
    batches: [],
    sheetsLink: 'https://docs.google.com/spreadsheets/d/1q2jA03C5yKXD8dHcGdECbZtGpUWOZDa1_YKQO8RkXjQ/edit?usp=sharing',
    generatedContent: ''
  });
  const [copiedEmailTemplate, setCopiedEmailTemplate] = useState<boolean>(false);
  // Subject lines for emails
  const [emailTemplateSubject, setEmailTemplateSubject] = useState<string>('');
  const [absentStudentEmailSubject, setAbsentStudentEmailSubject] = useState<string>('');
  const [presentStudentEmailSubject, setPresentStudentEmailSubject] = useState<string>('');
  const [copiedEmailTemplateSubject, setCopiedEmailTemplateSubject] = useState<boolean>(false);
  const [copiedAbsentStudentSubject, setCopiedAbsentStudentSubject] = useState<boolean>(false);
  const [copiedPresentStudentSubject, setCopiedPresentStudentSubject] = useState<boolean>(false);
  // Student email templates
  const [absentStudentEmailContent, setAbsentStudentEmailContent] = useState<string>('');
  const [presentStudentEmailContent, setPresentStudentEmailContent] = useState<string>('');
  const [copiedAbsentStudentEmail, setCopiedAbsentStudentEmail] = useState<boolean>(false);
  const [copiedPresentStudentEmail, setCopiedPresentStudentEmail] = useState<boolean>(false);
  const [selectedBatchForEmail, setSelectedBatchForEmail] = useState<string>('');

  // Security: Disable right-click and developer tools access
  useEffect(() => {
    const disableRightClick = (e: MouseEvent) => {
      e.preventDefault();
      return false;
    };

    const disableKeyboardShortcuts = (e: KeyboardEvent) => {
      // Disable F12 (Developer Tools)
      if (e.key === 'F12') {
        e.preventDefault();
        return false;
      }
      
      // Disable Ctrl+Shift+I (Developer Tools)
      if (e.ctrlKey && e.shiftKey && e.key === 'I') {
        e.preventDefault();
        return false;
      }
      
      // Disable Ctrl+Shift+J (Console)
      if (e.ctrlKey && e.shiftKey && e.key === 'J') {
        e.preventDefault();
        return false;
      }
      
      // Disable Ctrl+U (View Source)
      if (e.ctrlKey && e.key === 'u') {
        e.preventDefault();
        return false;
      }
      
      // Disable Ctrl+Shift+C (Element Inspector)
      if (e.ctrlKey && e.shiftKey && e.key === 'C') {
        e.preventDefault();
        return false;
      }
      
      // Disable Ctrl+S (Save Page)
      if (e.ctrlKey && e.key === 's') {
        e.preventDefault();
        return false;
      }
    };

    const disableTextSelection = () => {
      document.onselectstart = () => false;
      document.ondragstart = () => false;
    };

    const disablePrintScreen = (e: KeyboardEvent) => {
      // Disable Print Screen
      if (e.key === 'PrintScreen') {
        e.preventDefault();
        return false;
      }
    };

    // Add event listeners
    document.addEventListener('contextmenu', disableRightClick);
    document.addEventListener('keydown', disableKeyboardShortcuts);
    document.addEventListener('keydown', disablePrintScreen);
    
    // Disable text selection and drag
    disableTextSelection();
    
    // Disable developer tools detection (basic)
    const devToolsDetection = () => {
      if (window.outerHeight - window.innerHeight > 200 || window.outerWidth - window.innerWidth > 200) {
        document.body.innerHTML = '<div style="display:flex;justify-content:center;align-items:center;height:100vh;background:#000;color:#fff;font-family:Arial,sans-serif;">Access Denied</div>';
      }
    };
    
    // Check for developer tools periodically
    const detectionInterval = setInterval(devToolsDetection, 1000);

    // Cleanup event listeners
    return () => {
      document.removeEventListener('contextmenu', disableRightClick);
      document.removeEventListener('keydown', disableKeyboardShortcuts);
      document.removeEventListener('keydown', disablePrintScreen);
      document.onselectstart = null;
      document.ondragstart = null;
      clearInterval(detectionInterval);
    };
  }, []);

  const processExcelFile = (file: File, sheetNameToUse?: string) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      const workbook = XLSX.read(data, { type: 'array' });
      
      // Extract all sheet names
      const sheetNames = workbook.SheetNames;
      setAvailableSheets(sheetNames);
      
      // Use provided sheet name or default to first sheet
      const sheetName = sheetNameToUse || sheetNames[0];
      setSelectedSheet(sheetName);
      
      if (selectedCollege === 'niet') {
        // Process NIET attendance data for all sheets
        processAllSheetsForNIET(workbook, sheetName);
      } else {
        // Process regular student data for other colleges
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: ['A', 'B', 'C', 'D'] });
        
        const students: StudentData[] = (jsonData.slice(1) as Record<string, string | number>[]).map((row) => ({
          rank: Number(row.A) || 0,
          name: String(row.B) || '',
          rollNumber: String(row.C) || '',
          percentage: Number(row.D) || 0,
        })).filter(student => student.name && student.rollNumber);
        
        setStudentData(students);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const processAllSheetsForNIET = (workbook: XLSX.WorkBook, primarySheet: string) => {
    // Process the primary sheet first to get dates
    const primaryWorksheet = workbook.Sheets[primarySheet];
    const primaryRawData = XLSX.utils.sheet_to_json(primaryWorksheet, { header: 1, defval: '' }) as (string | number)[][];
    setRawAttendanceData(primaryRawData);
    
    // Extract dates from row 4 (index 3), starting from column H (index 7)
    const headerRow = primaryRawData[3] as string[];
    const dates: AttendanceDate[] = [];
    
    for (let i = 7; i < headerRow.length; i++) {
      const cellValue = headerRow[i];
      if (cellValue && typeof cellValue === 'string' && cellValue.trim()) {
        // Extract date from format like "28/07/2025 (9:10am to 01:10 pm)"
        const dateMatch = cellValue.match(/(\d{2}\/\d{2}\/\d{4})/);
        if (dateMatch) {
          dates.push({
            date: dateMatch[1],
            fullText: cellValue.trim(),
            columnIndex: i
          });
        }
      }
    }
    
    setAttendanceDates(dates);
    
    // If there are dates, select the first one by default but don't auto-process sheets
    if (dates.length > 0) {
      setSelectedDate(dates[0].date);
      calculateAttendanceStats(dates[0].date, dates[0].columnIndex, primaryRawData);
      
      // Don't auto-process sheets - wait for user selection
      // processSelectedSheetsAttendanceData will be called when user selects sheets
    }
  };

  const processSelectedSheetsAttendanceDataWithSet = (workbook: XLSX.WorkBook, selectedDate: string, sheetsToProcess: Set<string>) => {
    const allData = new Map<string, AttendanceStats>();
    
    // Only process sheets that are selected for processing
    sheetsToProcess.forEach(sheetName => {
      if (workbook.SheetNames.includes(sheetName)) {
        try {
          const worksheet = workbook.Sheets[sheetName];
          const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as (string | number)[][];
          
          // Find the correct column index for the selected date in this sheet
          const headerRow = rawData[3] as string[];
          let correctColumnIndex = -1;
          
          for (let i = 7; i < headerRow.length; i++) {
            const cellValue = headerRow[i];
            if (cellValue && typeof cellValue === 'string' && cellValue.trim()) {
              // Extract date from format like "28/07/2025 (9:10am to 01:10 pm)"
              const dateMatch = cellValue.match(/(\d{2}\/\d{2}\/\d{4})/);
              if (dateMatch && dateMatch[1] === selectedDate) {
                correctColumnIndex = i;
                break;
              }
            }
          }
          
          if (correctColumnIndex !== -1) {
            // Calculate attendance stats for this sheet with the correct column
            const stats = calculateAttendanceStatsForSheet(sheetName, selectedDate, correctColumnIndex, rawData);
            if (stats) {
              allData.set(sheetName, stats);
            }
          }
        } catch (error) {
          console.error(`Error processing sheet ${sheetName}:`, error);
        }
      }
    });
    
    setAllSheetsAttendanceData(allData);
  };

  const calculateAttendanceStatsForSheet = (sheetName: string, date: string, columnIndex: number, data: (string | number)[][]): AttendanceStats | null => {
    let present = 0;
    let absent = 0;
    const presentStudents: StudentAttendance[] = [];
    const absentStudents: StudentAttendance[] = [];
    
    // Get the maximum row to process based on the sheet type
    const maxRow = getMaxRowForSheet(sheetName);
    
    // Start from row 5 (index 4) to skip header rows, end before maxRow
    for (let i = 4; i < Math.min(data.length, maxRow); i++) {
      const row = data[i];
      if (row && row[columnIndex] !== undefined && row[columnIndex] !== '') {
        const attendanceValue = row[columnIndex];
        
        // Only process rows with valid attendance data (1 or 0)
        if (attendanceValue === 1 || attendanceValue === '1' || attendanceValue === 0 || attendanceValue === '0') {
          // Extract student name from column B (index 1) and email from column C (index 2)
          const studentName = row[1] ? String(row[1]).trim() : `Student ${i}`;
          let studentEmail = row[2] ? String(row[2]).trim() : '';
          
          // If email is not in proper format, try to construct it from the data
          if (!studentEmail.includes('@') || !studentEmail.endsWith('@niet.co.in')) {
            // Look for roll number or ID in adjacent columns to construct email
            const rollNumber = row[3] ? String(row[3]).trim() : row[2] ? String(row[2]).trim() : `student${i}`;
            studentEmail = `${rollNumber}`;
          }
          
          if (attendanceValue === 1 || attendanceValue === '1') {
            present++;
            presentStudents.push({
              name: studentName,
              email: studentEmail,
              status: 'present'
            });
          } else if (attendanceValue === 0 || attendanceValue === '0') {
            absent++;
            absentStudents.push({
              name: studentName,
              email: studentEmail,
              status: 'absent'
            });
          }
        }
      }
    }
    
    // Total students = present + absent (this ensures the math is always correct)
    const totalStudents = present + absent;
    
    if (totalStudents === 0) return null;
    
    return {
      date,
      totalStudents,
      present,
      absent,
      presentPercentage: Math.round((present / totalStudents) * 100),
      absentPercentage: Math.round((absent / totalStudents) * 100),
      presentStudents,
      absentStudents
    };
  };

  const getMaxRowForSheet = (sheetName: string): number => {
    const sheetNameLower = sheetName.toLowerCase();
    
    // Define row limits for different sheet types (subtract 1 for 0-based indexing)
    if (sheetNameLower.includes('ms-1') || sheetNameLower.includes('ms1')) {
      return 142 - 1; // Ignore row 142 and below, so process up to row 141 (index 141)
    }
    if (sheetNameLower.includes('java') && sheetNameLower.includes('sde-1')) {
      return 108 - 1; // Ignore row 108 and below, so process up to row 107 (index 107)
    }
    if (sheetNameLower.includes('java') && sheetNameLower.includes('sde-2')) {
      return 98 - 1; // Ignore row 98 and below, so process up to row 97 (index 97)
    }
    if (sheetNameLower.includes('data') && sheetNameLower.includes('scientist') && sheetNameLower.includes('python')) {
      return 112 - 1; // Ignore row 112 and below, so process up to row 111 (index 111)
    }
    if (sheetNameLower.includes('data') && sheetNameLower.includes('scientist') && sheetNameLower.includes('attendance')) {
      return 112 - 1; // Ignore row 112 and below, so process up to row 111 (index 111)
    }
    
    // Default: process all rows (no limit)
    return Number.MAX_SAFE_INTEGER;
  };

  const calculateAttendanceStats = (date: string, columnIndex: number, data: (string | number)[][]) => {
    let present = 0;
    let absent = 0;
    const presentStudents: StudentAttendance[] = [];
    const absentStudents: StudentAttendance[] = [];
    
    // Get the maximum row to process based on the selected sheet
    const maxRow = getMaxRowForSheet(selectedSheet);
    
    // Start from row 5 (index 4) to skip header rows, end before maxRow
    for (let i = 4; i < Math.min(data.length, maxRow); i++) {
      const row = data[i];
      if (row && row[columnIndex] !== undefined && row[columnIndex] !== '') {
        const attendanceValue = row[columnIndex];
        
        // Only process rows with valid attendance data (1 or 0)
        if (attendanceValue === 1 || attendanceValue === '1' || attendanceValue === 0 || attendanceValue === '0') {
          // Extract student name from column B (index 1) and email from column C (index 2)
          const studentName = row[1] ? String(row[1]).trim() : `Student ${i}`;
          let studentEmail = row[2] ? String(row[2]).trim() : '';
          
          // If email is not in proper format, try to construct it from the data
          if (!studentEmail.includes('@') || !studentEmail.endsWith('@niet.co.in')) {
            // Look for roll number or ID in adjacent columns to construct email
            const rollNumber = row[3] ? String(row[3]).trim() : row[2] ? String(row[2]).trim() : `student${i}`;
            studentEmail = `${rollNumber}`;
          }
          
          if (attendanceValue === 1 || attendanceValue === '1') {
            present++;
            presentStudents.push({
              name: studentName,
              email: studentEmail,
              status: 'present'
            });
          } else if (attendanceValue === 0 || attendanceValue === '0') {
            absent++;
            absentStudents.push({
              name: studentName,
              email: studentEmail,
              status: 'absent'
            });
          }
        }
      }
    }
    
    // Total students = present + absent (this ensures the math is always correct)
    const totalStudents = present + absent;
    
    const stats: AttendanceStats = {
      date,
      totalStudents,
      present,
      absent,
      presentPercentage: totalStudents > 0 ? Math.round((present / totalStudents) * 100) : 0,
      absentPercentage: totalStudents > 0 ? Math.round((absent / totalStudents) * 100) : 0,
      presentStudents,
      absentStudents
    };
    
    setAttendanceStats(stats);
  };

  const handleDateChange = (newDate: string) => {
    setSelectedDate(newDate);
    setCopiedPresentEmails(false);
    setCopiedAbsentEmails(false);
    const dateObj = attendanceDates.find(d => d.date === newDate);
    if (dateObj && rawAttendanceData.length > 0) {
      calculateAttendanceStats(newDate, dateObj.columnIndex, rawAttendanceData);
      
      // Also reprocess selected sheets for the new date if we have the file and selected sheets
      if (attendanceFile && selectedSheetsForProcessing.size > 0) {
        const reader = new FileReader();
        reader.onload = (e) => {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'array' });
          processSelectedSheetsAttendanceDataWithSet(workbook, newDate, selectedSheetsForProcessing);
        };
        reader.readAsArrayBuffer(attendanceFile);
      }
    }
    
    // Reset email template when date changes
    setEmailTemplate(prev => ({
      ...prev,
      trainingDate: newDate,
      generatedContent: ''
    }));
  };

  const formatEmailsForGmail = (students: StudentAttendance[]) => {
    // Filter out invalid emails and format for Gmail
    const validEmails = students
      .map(student => student.email)
      .filter(email => email.includes('@') && email.includes('.'))
      .map(email => email.trim());
    
    // Gmail accepts comma-separated emails, but we'll also ensure proper spacing
    return validEmails.join(', ');
  };

  const copyPresentEmails = async () => {
    if (!attendanceStats || attendanceStats.presentStudents.length === 0) return;
    
    try {
      const emailString = formatEmailsForGmail(attendanceStats.presentStudents);
      await navigator.clipboard.writeText(emailString);
      setCopiedPresentEmails(true);
      setTimeout(() => setCopiedPresentEmails(false), 3000);
    } catch (error) {
      console.error('Failed to copy present emails:', error);
      // Fallback for browsers that don't support clipboard API
      try {
        const emailString = formatEmailsForGmail(attendanceStats.presentStudents);
        const textArea = document.createElement('textarea');
        textArea.value = emailString;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        setCopiedPresentEmails(true);
        setTimeout(() => setCopiedPresentEmails(false), 3000);
      } catch (fallbackError) {
        console.error('Fallback copy also failed:', fallbackError);
      }
    }
  };

  const copyAbsentEmails = async () => {
    if (!attendanceStats || attendanceStats.absentStudents.length === 0) return;
    
    try {
      const emailString = formatEmailsForGmail(attendanceStats.absentStudents);
      await navigator.clipboard.writeText(emailString);
      setCopiedAbsentEmails(true);
      setTimeout(() => setCopiedAbsentEmails(false), 3000);
    } catch (error) {
      console.error('Failed to copy absent emails:', error);
      // Fallback for browsers that don't support clipboard API
      try {
        const emailString = formatEmailsForGmail(attendanceStats.absentStudents);
        const textArea = document.createElement('textarea');
        textArea.value = emailString;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        setCopiedAbsentEmails(true);
        setTimeout(() => setCopiedAbsentEmails(false), 3000);
      } catch (fallbackError) {
        console.error('Fallback copy also failed:', fallbackError);
      }
    }
  };

  const processInternReport = (reportText: string) => {
    if (!reportText.trim()) {
      setProcessedInternData([]);
      return;
    }

    // Basic processing - split by lines and filter meaningful data
    const lines = reportText.split('\n').filter(line => line.trim().length > 0);
    const processedData = lines.map((line, index) => ({
      id: index + 1,
      content: line.trim(),
      type: 'text'
    }));
    
    setProcessedInternData(processedData);
  };

  const handleSheetChange = (newSheetName: string) => {
    if (attendanceFile && newSheetName) {
      setSelectedSheet(newSheetName);
      // Reprocess the file with the new sheet
      processExcelFile(attendanceFile, newSheetName);
    }
  };

  const handleSheetSelectionToggle = (sheetName: string) => {
    const newSelection = new Set(selectedSheetsForProcessing);
    if (newSelection.has(sheetName)) {
      newSelection.delete(sheetName);
    } else {
      newSelection.add(sheetName);
    }
    setSelectedSheetsForProcessing(newSelection);
    
    // Reprocess attendance data for selected sheets if we have a selected date
    if (selectedDate && attendanceFile && newSelection.size > 0) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        processSelectedSheetsAttendanceDataWithSet(workbook, selectedDate, newSelection);
      };
      reader.readAsArrayBuffer(attendanceFile);
    } else if (newSelection.size === 0) {
      // Clear data if no sheets selected
      setAllSheetsAttendanceData(new Map());
    }
  };

  const selectAllSheets = () => {
    const allSheets = new Set(availableSheets);
    setSelectedSheetsForProcessing(allSheets);
    
    // Process all sheets if we have a selected date
    if (selectedDate && attendanceFile) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        processSelectedSheetsAttendanceDataWithSet(workbook, selectedDate, allSheets);
      };
      reader.readAsArrayBuffer(attendanceFile);
    }
  };

  const clearAllSheets = () => {
    setSelectedSheetsForProcessing(new Set());
    setAllSheetsAttendanceData(new Map());
  };

  const loadNIETAttendanceSheet = async () => {
    try {
      const response = await fetch('/Common Attendance Sheet data_NIET College.xlsx');
      const arrayBuffer = await response.arrayBuffer();
      const blob = new Blob([arrayBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const file = new File([blob], 'Common Attendance Sheet data_NIET College.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      setAttendanceFile(file);
      processExcelFile(file);
    } catch (error) {
      console.error('Failed to load NIET attendance sheet:', error);
    }
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files[0]) {
      const file = event.target.files[0];
      setAttendanceFile(file);
      processExcelFile(file);
    }
  };

  const toggleStudentSelection = (index: number) => {
    const newSelected = new Set(selectedStudents);
    if (newSelected.has(index)) {
      newSelected.delete(index);
    } else {
      newSelected.add(index);
    }
    setSelectedStudents(newSelected);
  };

  const generateEmails = () => {
    const selectedStudentData = studentData.filter((_, index) => selectedStudents.has(index));
    const emails = selectedStudentData.map(student => {
      const baseContent = `Dear ${student.name},

Greetings of the day!

We are pleased to share with you the results of your recent Daily Assessments till 18th August 2025 (15 sessions). Your average percentage is ${student.percentage}, reflecting your consistent effort and commitment.

We are also delighted to inform you that you have secured Rank ${student.rank} among your peers. Congratulations on this achievement!

Keep up the good work and continue striving for excellence.
With steady focus and dedication, you are sure to reach even greater milestones.

Wishing you all the best for your upcoming assessments.

If you have any questions feel free to reach out to ncetsupport@myanatomy.in

Warm regards,`;

      // Display version with HTML bold tags for visual rendering
      const displayVersion = baseContent
        .replace(`Your average percentage is ${student.percentage}`, `<b>Your average percentage is ${student.percentage}</b>`)
        .replace(`Rank ${student.rank}`, `<b>Rank ${student.rank}</b>`);

      // HTML version for rich text copying to email clients
      const htmlContent = baseContent
        .replace(`Your average percentage is ${student.percentage}`, `<strong>Your average percentage is ${student.percentage}</strong>`)
        .replace(`Rank ${student.rank}`, `<strong>Rank ${student.rank}</strong>`)
        .replace(/\n/g, '<br>');

      const htmlVersion = `<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; font-size: 14px; line-height: 1.4; color: #202124; background-color: #ffffff;">${htmlContent}</div>`;

      // Plain text version as fallback
      const plainTextVersion = baseContent;

      return { displayVersion, htmlVersion, plainTextVersion };
    });
    
    setGeneratedEmails(emails);
  };

  const copyToClipboard = async (emailData: EmailData, index: number) => {
    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([emailData.htmlVersion], { type: 'text/html' }),
          'text/plain': new Blob([emailData.plainTextVersion], { type: 'text/plain' }),
        })
      ];
      
      await navigator.clipboard.write(clipboardData);
      setCopiedEmailIndex(index);
      setTimeout(() => setCopiedEmailIndex(null), 2000);
    } catch (err) {
      console.error('Failed to copy text: ', err);
      // Fallback to plain text copy
      try {
        await navigator.clipboard.writeText(emailData.plainTextVersion);
        setCopiedEmailIndex(index);
        setTimeout(() => setCopiedEmailIndex(null), 2000);
      } catch (fallbackErr) {
        console.error('Fallback copy also failed: ', fallbackErr);
      }
    }
  };

  const copyAllEmails = async () => {
    try {
      const emailSeparator = '<div style="margin: 24px 0; padding: 12px 0; border-top: 1px solid #dadce0; color: #5f6368; font-style: italic; text-align: center;">--- Next Email ---</div>';
      const allEmailsHtml = generatedEmails.map(email => email.htmlVersion).join(emailSeparator);
      const allEmailsText = generatedEmails.map(email => email.plainTextVersion).join('\n\n--- Next Email ---\n\n');
      
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([allEmailsHtml], { type: 'text/html' }),
          'text/plain': new Blob([allEmailsText], { type: 'text/plain' }),
        })
      ];
      
      await navigator.clipboard.write(clipboardData);
      setCopiedEmailIndex(-1);
      setTimeout(() => setCopiedEmailIndex(null), 2000);
    } catch (err) {
      console.error('Failed to copy all emails: ', err);
      // Fallback to plain text copy
      try {
        const allEmailsText = generatedEmails.map(email => email.plainTextVersion).join('\n\n--- Next Email ---\n\n');
        await navigator.clipboard.writeText(allEmailsText);
        setCopiedEmailIndex(-1);
        setTimeout(() => setCopiedEmailIndex(null), 2000);
      } catch (fallbackErr) {
        console.error('Fallback copy also failed: ', fallbackErr);
      }
    }
  };

  const copySubjectLine = async () => {
    const subjectText = 'Congratulations on Your Daily Assessment Performance';
    try {
      await navigator.clipboard.writeText(subjectText);
      setCopiedSubject(true);
      setTimeout(() => setCopiedSubject(false), 2000);
    } catch (err) {
      console.error('Failed to copy subject line: ', err);
    }
  };

  const toggleEmailSent = (index: number) => {
    const newEmailsSent = new Set(emailsSent);
    if (newEmailsSent.has(index)) {
      newEmailsSent.delete(index);
    } else {
      newEmailsSent.add(index);
    }
    setEmailsSent(newEmailsSent);
  };

  const markAllEmailsSent = () => {
    const selectedIndices = Array.from(selectedStudents);
    setEmailsSent(new Set(selectedIndices));
  };

  const clearAllEmailsSent = () => {
    setEmailsSent(new Set());
  };

  // Collective student functions for all batches
  const getAllPresentStudents = (): StudentAttendance[] => {
    const allPresentStudents: StudentAttendance[] = [];
    
    for (const [, sheetStats] of allSheetsAttendanceData.entries()) {
      if (sheetStats.presentStudents) {
        allPresentStudents.push(...sheetStats.presentStudents);
      }
    }
    
    return allPresentStudents;
  };

  const getAllAbsentStudents = (): StudentAttendance[] => {
    const allAbsentStudents: StudentAttendance[] = [];
    
    for (const [, sheetStats] of allSheetsAttendanceData.entries()) {
      if (sheetStats.absentStudents) {
        allAbsentStudents.push(...sheetStats.absentStudents);
      }
    }
    
    return allAbsentStudents;
  };

  const copyAllPresentStudentEmails = async () => {
    const presentStudents = getAllPresentStudents();
    const emails = presentStudents.map(student => student.email).filter(email => email).join(', ');
    
    if (!emails) return;
    
    try {
      await navigator.clipboard.writeText(emails);
      setCopiedPresentEmails(true);
      setTimeout(() => setCopiedPresentEmails(false), 3000);
    } catch (error) {
      console.error('Failed to copy present student emails:', error);
    }
  };

  const copyAllAbsentStudentEmails = async () => {
    const absentStudents = getAllAbsentStudents();
    const emails = absentStudents.map(student => student.email).filter(email => email).join(', ');
    
    if (!emails) return;
    
    try {
      await navigator.clipboard.writeText(emails);
      setCopiedAbsentEmails(true);
      setTimeout(() => setCopiedAbsentEmails(false), 3000);
    } catch (error) {
      console.error('Failed to copy absent student emails:', error);
    }
  };

  const generateAllBatchDataFromAttendance = (): BatchData[] => {
    const batchData: BatchData[] = [];
    let batchId = 1;

    // Create batch data directly from the processed sheets data (allSheetsAttendanceData)
    // This ensures the email template uses EXACTLY the same data shown in the attendance analysis
    
    for (const [sheetName, sheetStats] of allSheetsAttendanceData.entries()) {
      // Determine batch name based on sheet name
      let batchName = sheetName; // Default to sheet name
      const sheetLower = sheetName.toLowerCase();
      
      if (sheetLower.includes('mern')) {
        batchName = 'MERN Stack Batch 1';
      } else if (sheetLower.includes('java') && sheetLower.includes('sde-1')) {
        batchName = 'Java SDE 1 Batch 2';
      } else if (sheetLower.includes('java') && sheetLower.includes('sde-2')) {
        batchName = 'Java SDE 2 Batch 3';
      } else if (sheetLower.includes('data') && sheetLower.includes('scientist')) {
        batchName = 'Data Science Python Batch 4';
      }

      batchData.push({
        id: batchId++,
        name: batchName,
        total: sheetStats.totalStudents,
        present: sheetStats.present,
        absent: sheetStats.absent
      });
    }

    // If no processed sheets data available, fall back to current attendance stats from the primary sheet
    if (batchData.length === 0 && attendanceStats && selectedSheet) {
      const sheetLower = selectedSheet.toLowerCase();
      let batchName = selectedSheet;
      
      if (sheetLower.includes('mern')) {
        batchName = 'MERN Stack Batch 1';
      } else if (sheetLower.includes('java') && sheetLower.includes('sde-1')) {
        batchName = 'Java SDE 1 Batch 2';
      } else if (sheetLower.includes('java') && sheetLower.includes('sde-2')) {
        batchName = 'Java SDE 2 Batch 3';
      } else if (sheetLower.includes('data') && sheetLower.includes('scientist')) {
        batchName = 'Data Science Python Batch 4';
      }

      batchData.push({
        id: 1,
        name: batchName,
        total: attendanceStats.totalStudents,
        present: attendanceStats.present,
        absent: attendanceStats.absent
      });
    }

    return batchData;
  };


  const generateEmailTemplate = () => {
    // Use actual attendance data from the analysis
    const actualDate = attendanceStats?.date || selectedDate || '21/08/2025';
    const actualBatches = generateAllBatchDataFromAttendance();
    
    // Format date for subject line (convert from DD/MM/YYYY to ordinal format)
    const formatDateForSubject = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('/');
      const dayNum = parseInt(day, 10);
      
      // Convert to ordinal
      let ordinalSuffix = 'th';
      if (dayNum % 100 < 11 || dayNum % 100 > 13) {
        switch (dayNum % 10) {
          case 1: ordinalSuffix = 'st'; break;
          case 2: ordinalSuffix = 'nd'; break;
          case 3: ordinalSuffix = 'rd'; break;
        }
      }
      
      const months = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      return `${dayNum}${ordinalSuffix} ${months[parseInt(month, 10)]} ${year}`;
    };
    
    const subjectLine = `NIET College NCET + Training Attendance for ${formatDateForSubject(actualDate)}`;
    
    // Apply topic mapping directly in email generation with batch numbers
    const getDisplayTopicName = (batchName: string): string => {
      if (batchName === 'MS-1' || batchName.toLowerCase().includes('ms-1')) {
        return 'MERN Stack Batch 1';
      } else if (batchName === 'Java SDE-1' || batchName.toLowerCase().includes('java sde-1') || batchName.toLowerCase().includes('java sde 1')) {
        return 'Java SDE Batch 2';
      } else if (batchName === 'Java SDE-2' || batchName.toLowerCase().includes('java sde-2') || batchName.toLowerCase().includes('java sde 2')) {
        return 'Java SDE Batch 3';
      } else if (batchName.toLowerCase().includes('data scientist') || batchName.toLowerCase().includes('data science python')) {
        return 'Data Science Python Batch 4';
      }
      return batchName; // fallback to original name
    };
    
    let batchContent = '';
    if (actualBatches.length > 0) {
      batchContent = actualBatches.map(batch => {
        const displayName = getDisplayTopicName(batch.name);
        const absentFormatted = String(batch.absent).padStart(2, '0');
        return `        <strong>·        Total Number of Registered Students for ${displayName}: ${batch.total}</strong><br>
        <strong>·        Number of Students Present for ${displayName}: ${batch.present}</strong><br>
        <strong>·        Number of Students Absent for ${displayName}: ${absentFormatted}</strong>`;
      }).join('<br><br>');
    } else {
      batchContent = 'No attendance data available. Please select sheets to process and ensure you have selected a date in the Attendance Analysis section.';
    }
    
    const htmlContent = `<div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
<p>Dear Ma'am/Sir,</p>

<p>Greetings of the day!</p>

<p>I hope you are doing well.</p>

<p>This is to inform you that the training session conducted on ${actualDate} for the <strong>NCET + Training</strong> was successfully completed. Please find below the attendance details of the students who participated in the session:</p>

<div style="margin-left: 20px;">
${batchContent}
</div>

<p>I'm sharing the sheet attached in the email for your reference. It includes <strong>detailed attendance</strong>, <strong>daily Assessment</strong>, <strong>and a list of absent students</strong> during each training session.</p>

<p><strong>Link for Daily Attendance and Assessment: </strong> <a href="${emailTemplate.sheetsLink}">${emailTemplate.sheetsLink}</a></p>

<p>Kindly go through the same and let us know if you have any questions or need any further information.</p>

<p>Thank you for your continued support and coordination.</p>

<p><strong>Regards,</strong></p>
</div>`;

    // Plain text version for fallback
    const plainTextContent = `Dear Ma'am/Sir,

Greetings of the day!

I hope you are doing well.

This is to inform you that the training session conducted on ${actualDate} for the NCET + Training was successfully completed. Please find below the attendance details of the students who participated in the session:

${actualBatches.length > 0 ? actualBatches.map(batch => {
  const displayName = getDisplayTopicName(batch.name);
  const absentFormatted = String(batch.absent).padStart(2, '0');
  return `        ·        Total Number of Registered Students for ${displayName}: ${batch.total}
        ·        Number of Students Present for ${displayName}: ${batch.present}
        ·        Number of Students Absent for ${displayName}: ${absentFormatted}`;
}).join('\n\n') : 'No attendance data available. Please select sheets to process and ensure you have selected a date in the Attendance Analysis section.'}

I'm sharing the sheet attached in the email for your reference. It includes detailed attendance, daily Assessment, and a list of absent students during each training session.

Link for Daily Attendance and Assessment: ${emailTemplate.sheetsLink}

Kindly go through the same and let us know if you have any questions or need any further information.

Thank you for your continued support and coordination.

Regards,`;

    // Update email template with actual data
    setEmailTemplate(prev => ({ 
      ...prev, 
      trainingDate: actualDate,
      batches: actualBatches,
      generatedContent: htmlContent,
      plainTextContent: plainTextContent
    }));
    
    // Set the subject line
    setEmailTemplateSubject(subjectLine);
  };

  const copyEmailTemplate = async () => {
    if (!emailTemplate.generatedContent) {
      generateEmailTemplate();
      return;
    }

    try {
      // Try to copy both HTML and plain text versions for better Gmail compatibility
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([emailTemplate.generatedContent], { type: 'text/html' }),
          'text/plain': new Blob([emailTemplate.plainTextContent || emailTemplate.generatedContent], { type: 'text/plain' }),
        })
      ];
      
      await navigator.clipboard.write(clipboardData);
      setCopiedEmailTemplate(true);
      setTimeout(() => setCopiedEmailTemplate(false), 3000);
    } catch (error) {
      console.error('Failed to copy email template with formatting:', error);
      // Fallback: try copying just HTML content
      try {
        await navigator.clipboard.writeText(emailTemplate.generatedContent);
        setCopiedEmailTemplate(true);
        setTimeout(() => setCopiedEmailTemplate(false), 3000);
      } catch (fallbackError) {
        console.error('Fallback copy also failed:', fallbackError);
        // Last resort: manual copy with textarea
        try {
          const textArea = document.createElement('textarea');
          textArea.value = emailTemplate.plainTextContent || emailTemplate.generatedContent;
          document.body.appendChild(textArea);
          textArea.select();
          document.execCommand('copy');
          document.body.removeChild(textArea);
          setCopiedEmailTemplate(true);
          setTimeout(() => setCopiedEmailTemplate(false), 3000);
        } catch (finalError) {
          console.error('All copy methods failed:', finalError);
        }
      }
    }
  };


  const generateTopicsTableFromInternReport = () => {
    if (!processedInternData || processedInternData.length === 0) {
      return '<tr><td colspan="3" style="text-align: center; color: #666;">No intern report data available</td></tr>';
    }

    if (selectedBatchForEmail === '') {
      return '<tr><td colspan="3" style="text-align: center; color: #666;">Please select a batch for the email</td></tr>';
    }

    // Map batch names to proper topic names (handle names with "Attendance" suffix)
    const getTopicName = (batchName: string): string => {
      const normalizedName = batchName.toLowerCase();
      
      if (normalizedName.includes('ms-1')) {
        return 'MERN Stack';
      } else if (normalizedName.includes('java sde-1')) {
        return 'Java SDE';
      } else if (normalizedName.includes('java sde-2')) {
        return 'Java SDE';
      } else if (normalizedName.includes('data scientist')) {
        return 'Data Science Python';
      }
      return batchName; // fallback to original name
    };

    const topicName = getTopicName(selectedBatchForEmail);
    
    // Create numbered points from intern report data (each line becomes a numbered point)
    const numberedDescription = processedInternData.map((item, index) => 
      `${index + 1}. ${item.content}`
    ).join('<br/>');
    
    return `<tr>
      <td style="padding: 8px; border: 1px solid #ccc; text-align: center;">1</td>
      <td style="padding: 8px; border: 1px solid #ccc;"><strong>${topicName}</strong></td>
      <td style="padding: 8px; border: 1px solid #ccc;">${numberedDescription}</td>
    </tr>`;
  };

  const generateAbsentStudentEmail = () => {
    const actualDate = attendanceStats?.date || selectedDate || '21/08/2025';
    const topicsTable = generateTopicsTableFromInternReport();
    
    // Format date for subject line (convert from DD/MM/YYYY to DD Month YYYY)
    const formatDateForStudentSubject = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('/');
      const dayNum = parseInt(day, 10);
      const months = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      return `${dayNum} ${months[parseInt(month, 10)]} ${year}`;
    };
    
    // Get topic name for subject (handle names with "Attendance" suffix)
    const getTopicName = (batchName: string): string => {
      const normalizedName = batchName.toLowerCase();
      
      if (normalizedName.includes('ms-1')) {
        return 'MERN Stack';
      } else if (normalizedName.includes('java sde-1') || normalizedName.includes('java sde-2')) {
        return 'Java SDE';
      } else if (normalizedName.includes('data scientist')) {
        return 'Data Science Python';
      }
      return batchName; // fallback to original name
    };
    
    const subjectTopic = selectedBatchForEmail ? getTopicName(selectedBatchForEmail) : 'Training';
    const subjectLine = `${subjectTopic} NCET + Training NIET College Attendance ${formatDateForStudentSubject(actualDate)}`;
    
    const htmlContent = `<div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
<p><strong>Dear Students,</strong></p>

<p>This email is to address the issue of student attendance at our live training sessions. We have observed that some students have missed the live training session on <strong>${actualDate}</strong>.</p>

<p><strong>Here's what you missed during the session:</strong></p>

<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
  <thead>
    <tr style="background-color: #f0f0f0;">
      <th style="padding: 10px; border: 1px solid #ccc; text-align: center;"><strong>S.No</strong></th>
      <th style="padding: 10px; border: 1px solid #ccc;"><strong>Topic</strong></th>
      <th style="padding: 10px; border: 1px solid #ccc;"><strong>Description</strong></th>
    </tr>
  </thead>
  <tbody>
    ${topicsTable}
  </tbody>
</table>

<p>We understand that unforeseen circumstances may arise, however, it is crucial to attend these sessions regularly. These live sessions are an integral part of your learning journey and provide valuable opportunities for interactive learning, Q&A, and engagement with instructors and fellow students.</p>

<p>Missing these free sessions is not only detrimental to your learning but also disrespectful to the instructors and other students who are diligently participating.</p>

<p>Students who continue to remain absent for sessions will be flagged, and appropriate escalations will be made with the Training and Placement Officers (TPOs) if this behaviour is continued.</p>

<p>We expect all students to attend all upcoming live training sessions promptly.</p>

<p>We urge you to prioritize your attendance and actively participate in these valuable sessions.</p>

<p><strong>Regards,</strong></p>
</div>`;

    setAbsentStudentEmailContent(htmlContent);
    setAbsentStudentEmailSubject(subjectLine);
  };

  const generatePresentStudentEmail = () => {
    const actualDate = attendanceStats?.date || selectedDate || '21/08/2025';
    const topicsTable = generateTopicsTableFromInternReport();
    
    // Format date for subject line (convert from DD/MM/YYYY to DD Month YYYY)
    const formatDateForStudentSubject = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('/');
      const dayNum = parseInt(day, 10);
      const months = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      return `${dayNum} ${months[parseInt(month, 10)]} ${year}`;
    };
    
    // Get topic name for subject (handle names with "Attendance" suffix)
    const getTopicName = (batchName: string): string => {
      const normalizedName = batchName.toLowerCase();
      
      if (normalizedName.includes('ms-1')) {
        return 'MERN Stack';
      } else if (normalizedName.includes('java sde-1') || normalizedName.includes('java sde-2')) {
        return 'Java SDE';
      } else if (normalizedName.includes('data scientist')) {
        return 'Data Science Python';
      }
      return batchName; // fallback to original name
    };
    
    const subjectTopic = selectedBatchForEmail ? getTopicName(selectedBatchForEmail) : 'Training';
    const subjectLine = `${subjectTopic} NCET + Training NIET College Attendance ${formatDateForStudentSubject(actualDate)}`;
    
    const htmlContent = `<div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
<p><strong>Dear Students,</strong></p>

<p>On behalf of the <strong>NCET Live Training</strong> team, we would like to congratulate you on your punctuality in attending the recent live training session conducted on <strong>${actualDate}</strong>.</p>

<p><strong>We appreciate your dedication and commitment to learning.</strong></p>

<p><strong>Here's a quick recap of what was discussed during the session:</strong></p>

<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
  <thead>
    <tr style="background-color: #f0f0f0;">
      <th style="padding: 10px; border: 1px solid #ccc; text-align: center;"><strong>S.No</strong></th>
      <th style="padding: 10px; border: 1px solid #ccc;"><strong>Topic</strong></th>
      <th style="padding: 10px; border: 1px solid #ccc;"><strong>Description</strong></th>
    </tr>
  </thead>
  <tbody>
    ${topicsTable}
  </tbody>
</table>

<p>We have also received feedback from many of you regarding the sessions and go through it continuously to identify how we can improve the process.</p>

<p><strong>Thank you for your continued support and participation.</strong></p>

<p><strong>Regards,</strong></p>
</div>`;

    setPresentStudentEmailContent(htmlContent);
    setPresentStudentEmailSubject(subjectLine);
  };

  const copyAbsentStudentEmail = async () => {
    if (!absentStudentEmailContent) {
      generateAbsentStudentEmail();
      return;
    }

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([absentStudentEmailContent], { type: 'text/html' }),
          'text/plain': new Blob([absentStudentEmailContent.replace(/<[^>]*>/g, '')], { type: 'text/plain' }),
        })
      ];
      
      await navigator.clipboard.write(clipboardData);
      setCopiedAbsentStudentEmail(true);
      setTimeout(() => setCopiedAbsentStudentEmail(false), 3000);
    } catch (error) {
      console.error('Failed to copy absent student email:', error);
    }
  };

  const copyPresentStudentEmail = async () => {
    if (!presentStudentEmailContent) {
      generatePresentStudentEmail();
      return;
    }

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([presentStudentEmailContent], { type: 'text/html' }),
          'text/plain': new Blob([presentStudentEmailContent.replace(/<[^>]*>/g, '')], { type: 'text/plain' }),
        })
      ];
      
      await navigator.clipboard.write(clipboardData);
      setCopiedPresentStudentEmail(true);
      setTimeout(() => setCopiedPresentStudentEmail(false), 3000);
    } catch (error) {
      console.error('Failed to copy present student email:', error);
    }
  };

  // Copy functions for subject lines
  const copyEmailTemplateSubject = async () => {
    if (!emailTemplateSubject) return;
    
    try {
      await navigator.clipboard.writeText(emailTemplateSubject);
      setCopiedEmailTemplateSubject(true);
      setTimeout(() => setCopiedEmailTemplateSubject(false), 3000);
    } catch (error) {
      console.error('Failed to copy email template subject:', error);
    }
  };

  const copyAbsentStudentEmailSubject = async () => {
    if (!absentStudentEmailSubject) return;
    
    try {
      await navigator.clipboard.writeText(absentStudentEmailSubject);
      setCopiedAbsentStudentSubject(true);
      setTimeout(() => setCopiedAbsentStudentSubject(false), 3000);
    } catch (error) {
      console.error('Failed to copy absent student email subject:', error);
    }
  };

  const copyPresentStudentEmailSubject = async () => {
    if (!presentStudentEmailSubject) return;
    
    try {
      await navigator.clipboard.writeText(presentStudentEmailSubject);
      setCopiedPresentStudentSubject(true);
      setTimeout(() => setCopiedPresentStudentSubject(false), 3000);
    } catch (error) {
      console.error('Failed to copy present student email subject:', error);
    }
  };

  const isUploadComplete = attendanceFile !== null;

  return (
    <div className="min-h-screen bg-gray-900 text-gray-200 font-sans p-4 sm:p-6 lg:p-8 select-none" style={{ userSelect: 'none', WebkitUserSelect: 'none', msUserSelect: 'none', WebkitTouchCallout: 'none' }}>
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <header className="flex items-center justify-between mb-8">
          <div className="flex items-center gap-4">
            <Mail className="w-8 h-8 text-gray-400" />
            <div>
              <h1 className="text-2xl font-bold text-white">Attendance Compilation</h1>
              <p className="text-gray-400">Training attendance emails</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-400">College:</span>
              <div className="relative">
                <select
                  value={selectedCollege}
                  onChange={(e) => {
                    setSelectedCollege(e.target.value);
                    // Reset all data when college changes
                    setAttendanceFile(null);
                    setStudentData([]);
                    setSelectedStudents(new Set());
                    setGeneratedEmails([]);
                    setEmailsSent(new Set());
                    // Reset NIET-specific data
                    setAttendanceDates([]);
                    setSelectedDate('');
                    setAttendanceStats(null);
                    setRawAttendanceData([]);
                    setCopiedPresentEmails(false);
                    setCopiedAbsentEmails(false);
                    // Reset intern report data
                    setInternReport('');
                    setProcessedInternData([]);
                    setInternReportExpanded(false);
                    // Reset sheet selection data
                    setAvailableSheets([]);
                    setSelectedSheet('');
                    setSelectedSheetsForProcessing(new Set());
                    setAllSheetsAttendanceData(new Map());
                  }}
                  className="flex items-center bg-gray-800 border border-gray-700 rounded-md px-3 py-1.5 text-sm font-medium text-white cursor-pointer focus:ring-2 focus:ring-blue-500 focus:border-blue-500 appearance-none pr-8 select-auto"
                  style={{ userSelect: 'auto', WebkitUserSelect: 'auto' }}
                >
                  <option value="" disabled className="text-gray-400">Select College</option>
                  {colleges.map((college) => (
                    <option key={college.id} value={college.id} className="text-white bg-gray-800">
                      {college.name}
                    </option>
                  ))}
                </select>
                <div className="absolute inset-y-0 right-0 flex items-center pr-2 pointer-events-none">
                  <Building className="w-4 h-4 mr-1 text-gray-400" />
                  <ChevronDown className="w-4 h-4 text-gray-500" />
                </div>
              </div>
            </div>
            <button className="p-2 rounded-full hover:bg-gray-800">
              <Moon className="w-5 h-5 text-gray-400" />
            </button>
          </div>
        </header>

        <main className="space-y-8">
          {/* Row 1: Upload and Student Selection */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Box 1: Upload Document Section */}
            <section className={`bg-gray-800/50 border border-gray-700/50 rounded-xl p-6 transition-opacity ${!selectedCollege ? 'opacity-50' : 'opacity-100'}`}>
            <div className="flex items-center gap-3 mb-4">
              <Upload className="w-5 h-5 text-gray-400" />
              <h2 className="text-lg font-semibold text-white">Upload Files</h2>
              {selectedCollege && (
                <span className="bg-blue-600 text-white text-xs px-2 py-1 rounded-full">
                  {colleges.find(c => c.id === selectedCollege)?.name}
                </span>
              )}
            </div>
            <label htmlFor="attendance-sheet" className="text-sm text-gray-400 mb-2 block">Attendance Sheet</label>
            {!selectedCollege ? (
              <div className="relative border-2 border-dashed border-gray-600 rounded-lg p-10 text-center">
                <div className="flex flex-col items-center justify-center">
                  <Building className="w-10 h-10 text-gray-600 mb-3" />
                  <p className="text-gray-500 font-medium">Please select a college first</p>
                  <p className="text-xs text-gray-600 mt-1">Choose from SRM, Karpagam, or NIET above</p>
                </div>
              </div>
            ) : selectedCollege === 'niet' ? (
              <div className="space-y-4">
                <div className="relative border-2 border-dashed border-blue-600 rounded-lg p-8 text-center bg-blue-600/10">
                  <div className="flex flex-col items-center justify-center">
                    <FileText className="w-10 h-10 text-blue-500 mb-3" />
                    <p className="text-white font-medium mb-2">NIET Attendance Sheet</p>
                    <p className="text-xs text-gray-400 mb-4">
                      {attendanceFile ? `Loaded: ${attendanceFile.name}` : 'Use the pre-configured attendance sheet or upload a custom one'}
                    </p>
                    <button 
                      onClick={loadNIETAttendanceSheet}
                      className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors font-medium"
                    >
                      Load NIET Attendance Sheet
                    </button>
                  </div>
                </div>
                
                <div className="relative">
                  <div className="absolute inset-0 flex items-center">
                    <div className="w-full border-t border-gray-600"></div>
                  </div>
                  <div className="relative flex justify-center text-xs uppercase">
                    <span className="bg-gray-800 px-2 text-gray-400">Or upload custom file</span>
                  </div>
                </div>
                
                <div className="relative border-2 border-dashed border-gray-600 rounded-lg p-6 text-center cursor-pointer hover:border-gray-500 transition-colors">
                  <input 
                    id="attendance-sheet-custom" 
                    type="file" 
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    onChange={handleFileChange}
                    accept=".xlsx, .xls, .csv"
                    disabled={!selectedCollege}
                  />
                  <div className="flex flex-col items-center justify-center">
                    <Upload className="w-6 h-6 text-gray-500 mb-2" />
                    <p className="text-sm text-gray-300">Upload custom attendance sheet</p>
                  </div>
                </div>
              </div>
            ) : (
              <div className="relative border-2 border-dashed border-gray-600 rounded-lg p-10 text-center cursor-pointer hover:border-gray-500 transition-colors">
                <input 
                  id="attendance-sheet" 
                  type="file" 
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  onChange={handleFileChange}
                  accept=".xlsx, .xls, .csv"
                  disabled={!selectedCollege}
                />
                <div className="flex flex-col items-center justify-center">
                  <FileText className="w-10 h-10 text-gray-500 mb-3" />
                  <p className="text-white font-medium">
                    {attendanceFile ? `File selected: ${attendanceFile.name}` : "Click to browse or drag file here"}
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    Excel file with student data (A=Rank, B=Name, C=Roll Number, D=Percentage)
                  </p>
                  <p className="text-xs text-gray-500 mt-1">Maximum file size: 10 MB</p>
                </div>
              </div>
            )}
            {/* Sheet Selection */}
            {availableSheets.length > 1 && (
              <div className="mt-4 p-4 bg-blue-600/10 border border-blue-500/30 rounded-lg">
                <div className="flex items-center justify-between mb-3">
                  <div>
                    <h3 className="text-sm font-semibold text-blue-300 mb-1">Excel Sheet Selection</h3>
                    <p className="text-xs text-blue-200/80">Multiple sheets detected. Primary sheet for attendance analysis:</p>
                  </div>
                  <span className="bg-blue-600 text-white text-xs px-2 py-1 rounded-full">
                    {availableSheets.length} sheets
                  </span>
                </div>
                
                {/* Primary Sheet Selection */}
                <div className="mb-4">
                  <label className="text-xs text-blue-200 font-medium mb-2 block">Primary Sheet (for date selection):</label>
                  <select
                    value={selectedSheet}
                    onChange={(e) => handleSheetChange(e.target.value)}
                    className="w-full bg-gray-700 border border-gray-600 rounded-md px-3 py-2 text-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 select-auto"
                    style={{ userSelect: 'auto', WebkitUserSelect: 'auto' }}
                  >
                    {availableSheets.map((sheetName) => (
                      <option key={sheetName} value={sheetName} className="text-white bg-gray-700">
                        📊 {sheetName}
                        {getMaxRowForSheet(sheetName) !== Number.MAX_SAFE_INTEGER && 
                          ` (rows up to ${getMaxRowForSheet(sheetName)})`
                        }
                      </option>
                    ))}
                  </select>
                </div>
                
                {/* Manual Sheet Selection */}
                <div className="mt-4 p-4 bg-gray-700/30 border border-gray-600/50 rounded-lg">
                  <div className="flex items-center justify-between mb-3">
                    <h4 className="text-sm font-semibold text-white">Select Sheets to Process</h4>
                    <div className="flex gap-2">
                      <button
                        onClick={selectAllSheets}
                        className="text-xs px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded transition-colors"
                      >
                        Select All
                      </button>
                      <button
                        onClick={clearAllSheets}
                        className="text-xs px-2 py-1 bg-gray-600 hover:bg-gray-700 text-white rounded transition-colors"
                      >
                        Clear All
                      </button>
                    </div>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-48 overflow-y-auto">
                    {availableSheets.map((sheetName) => (
                      <label
                        key={sheetName}
                        className="flex items-center gap-3 p-2 hover:bg-gray-600/50 rounded cursor-pointer"
                      >
                        <input
                          type="checkbox"
                          checked={selectedSheetsForProcessing.has(sheetName)}
                          onChange={() => handleSheetSelectionToggle(sheetName)}
                          className="w-4 h-4 text-blue-600 bg-gray-700 border-gray-600 rounded focus:ring-blue-500"
                        />
                        <div className="flex-1">
                          <div className="text-sm text-white font-medium">📊 {sheetName}</div>
                          {getMaxRowForSheet(sheetName) !== Number.MAX_SAFE_INTEGER && (
                            <div className="text-xs text-orange-400">
                              Rows up to {getMaxRowForSheet(sheetName)}
                            </div>
                          )}
                        </div>
                      </label>
                    ))}
                  </div>
                  
                  <div className="mt-3 p-2 bg-blue-600/10 border border-blue-500/30 rounded text-center">
                    <div className="text-sm font-medium text-blue-300">
                      {selectedSheetsForProcessing.size} of {availableSheets.length} sheets selected
                    </div>
                    <div className="text-xs text-blue-200/80 mt-1">
                      Processed: {allSheetsAttendanceData.size} sheets with data
                    </div>
                  </div>
                </div>
                
                {selectedSheet && getMaxRowForSheet(selectedSheet) !== Number.MAX_SAFE_INTEGER && (
                  <div className="mt-2 p-2 bg-orange-600/10 border border-orange-500/30 rounded text-xs text-orange-300">
                    ⚠️ <strong>Row Filtering Active:</strong> Processing data up to row {getMaxRowForSheet(selectedSheet)} for {selectedSheet} sheet type
                  </div>
                )}
              </div>
            )}

            {/* Single Sheet Info */}
            {availableSheets.length === 1 && selectedSheet && (
              <div className="mt-4 p-3 bg-gray-700/30 border border-gray-600/50 rounded-lg">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-gray-400">Using Excel Sheet:</span>
                    <span className="text-sm font-medium text-white">📊 {selectedSheet}</span>
                  </div>
                  {getMaxRowForSheet(selectedSheet) !== Number.MAX_SAFE_INTEGER && (
                    <div className="text-xs text-orange-400">
                      ⚠️ Processing rows up to {getMaxRowForSheet(selectedSheet)}
                    </div>
                  )}
                </div>
              </div>
            )}

            {(studentData.length > 0 || attendanceDates.length > 0) && (
              <div className="mt-4 p-4 bg-gray-700/50 rounded-lg">
                <h3 className="text-sm font-medium text-white mb-2">
                  {selectedCollege === 'niet' 
                    ? `Attendance Data Loaded (${attendanceDates.length} dates found)`
                    : `Processed Student Data (${studentData.length} students)`
                  }
                  {selectedSheet && (
                    <span className="ml-2 text-xs bg-gray-600 text-gray-300 px-2 py-0.5 rounded">
                      Sheet: {selectedSheet}
                    </span>
                  )}
                </h3>
                <div className="max-h-32 overflow-y-auto">
                  {selectedCollege === 'niet' ? (
                    attendanceDates.slice(0, 5).map((dateObj, index) => (
                      <div key={index} className="text-xs text-gray-300 py-1">
                        {dateObj.date} - {dateObj.fullText}
                      </div>
                    ))
                  ) : (
                    studentData.slice(0, 5).map((student, index) => (
                      <div key={index} className="text-xs text-gray-300 py-1">
                        {student.rank}. {student.name} ({student.rollNumber}) - {student.percentage}
                      </div>
                    ))
                  )}
                  {selectedCollege === 'niet' ? (
                    attendanceDates.length > 5 && (
                      <div className="text-xs text-gray-400">...and {attendanceDates.length - 5} more dates</div>
                    )
                  ) : (
                    studentData.length > 5 && (
                      <div className="text-xs text-gray-400">...and {studentData.length - 5} more students</div>
                    )
                  )}
                </div>
              </div>
            )}
            </section>

            {/* Box 2: Student Selection */}
            <section className={`bg-gray-800/50 border border-gray-700/50 rounded-xl p-6 transition-opacity ${!selectedCollege || !isUploadComplete ? 'opacity-50' : 'opacity-100'}`}>
            <div className="flex items-center gap-3 mb-4">
              <Book className="w-5 h-5 text-gray-400" />
              <h2 className="text-lg font-semibold text-white">
                {selectedCollege === 'niet' ? 'Attendance Analysis' : 'Select Students'}
              </h2>
            </div>
            <p className="text-sm text-gray-400 mb-4">
              {selectedCollege === 'niet' 
                ? 'Select a date to view attendance statistics.'
                : 'Choose which students you want to send emails to.'
              }
            </p>
            {selectedCollege && isUploadComplete && (
              (selectedCollege === 'niet' && attendanceDates.length > 0) || 
              (selectedCollege !== 'niet' && studentData.length > 0)
            ) ? (
              selectedCollege === 'niet' ? (
                // NIET Attendance Analysis
                <div className="space-y-4">
                  <div className="flex justify-between items-center">
                    <span className="text-sm text-gray-300">
                      {attendanceDates.length} attendance dates available
                    </span>
                    <div className="flex gap-2">
                      <span className="text-xs px-2 py-1 bg-blue-600 text-white rounded">
                        NIET College
                      </span>
                      <span className="text-xs px-2 py-1 bg-green-600 text-white rounded">
                        {selectedSheetsForProcessing.size} selected | {allSheetsAttendanceData.size} processed
                      </span>
                    </div>
                  </div>
                  
                  {/* Date Selection */}
                  <div className="space-y-3">
                    <label className="text-sm text-gray-300 font-medium">Select Date:</label>
                    <select
                      value={selectedDate}
                      onChange={(e) => handleDateChange(e.target.value)}
                      className="w-full bg-gray-700 border border-gray-600 rounded-md px-3 py-2 text-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 select-auto"
                      style={{ userSelect: 'auto', WebkitUserSelect: 'auto' }}
                    >
                      {attendanceDates.map((dateObj) => (
                        <option key={dateObj.date} value={dateObj.date}>
                          {dateObj.date} - {dateObj.fullText}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Attendance Statistics */}
                  {attendanceStats && (
                    <div className="bg-gray-700/50 rounded-lg p-4">
                      <h3 className="text-lg font-semibold text-white mb-3">
                        Attendance for {attendanceStats.date}
                      </h3>
                      
                      <div className="grid grid-cols-2 gap-4 mb-4">
                        <div className="bg-green-600/20 border border-green-500/30 rounded-lg p-3">
                          <div className="text-green-300 text-sm font-medium">Present</div>
                          <div className="text-2xl font-bold text-white">{attendanceStats.present}</div>
                          <div className="text-xs text-green-400">{attendanceStats.presentPercentage}%</div>
                        </div>
                        
                        <div className="bg-red-600/20 border border-red-500/30 rounded-lg p-3">
                          <div className="text-red-300 text-sm font-medium">Absent</div>
                          <div className="text-2xl font-bold text-white">{attendanceStats.absent}</div>
                          <div className="text-xs text-red-400">{attendanceStats.absentPercentage}%</div>
                        </div>
                      </div>
                      
                      <div className="bg-blue-600/20 border border-blue-500/30 rounded-lg p-3">
                        <div className="text-blue-300 text-sm font-medium">Total Students</div>
                        <div className="text-xl font-bold text-white">{attendanceStats.totalStudents}</div>
                      </div>
                      
                      {/* Progress Bar */}
                      <div className="mt-4">
                        <div className="flex justify-between text-xs text-gray-400 mb-1">
                          <span>Present: {attendanceStats.presentPercentage}%</span>
                          <span>Absent: {attendanceStats.absentPercentage}%</span>
                        </div>
                        <div className="w-full bg-gray-700 rounded-full h-2">
                          <div 
                            className="bg-green-600 h-2 rounded-full transition-all duration-300"
                            style={{ width: `${attendanceStats.presentPercentage}%` }}
                          ></div>
                        </div>
                      </div>
                      
                      {/* All Sheets Summary */}
                      {allSheetsAttendanceData.size > 0 && (
                        <div className="mt-6 bg-gray-600/30 rounded-lg p-4">
                          <h4 className="text-sm font-semibold text-white mb-3">
                            All Sheets Summary for {attendanceStats.date}
                          </h4>
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                            {Array.from(allSheetsAttendanceData.entries()).map(([sheetName, stats]) => (
                              <div key={sheetName} className="bg-gray-700/50 rounded p-3">
                                <div className="text-xs font-medium text-blue-300 mb-2">{sheetName}</div>
                                <div className="grid grid-cols-3 gap-2 text-xs">
                                  <div className="text-center">
                                    <div className="text-gray-400">Total</div>
                                    <div className="text-white font-bold">{stats.totalStudents}</div>
                                  </div>
                                  <div className="text-center">
                                    <div className="text-gray-400">Present</div>
                                    <div className="text-green-400 font-bold">{stats.present}</div>
                                  </div>
                                  <div className="text-center">
                                    <div className="text-gray-400">Absent</div>
                                    <div className="text-red-400 font-bold">{stats.absent}</div>
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                      
                      {/* Gmail Integration Instructions */}
                      <div className="mt-6 bg-blue-600/10 border border-blue-500/30 rounded-lg p-4">
                        <div className="flex items-start gap-3">
                          <Mail className="w-5 h-5 text-blue-400 mt-0.5 flex-shrink-0" />
                          <div>
                            <h4 className="text-blue-300 text-sm font-semibold mb-2">How to Send Emails via Gmail</h4>
                            <ol className="text-xs text-blue-200/80 space-y-1 list-decimal list-inside">
                              <li>Click &ldquo;Copy for Gmail&rdquo; button for Present or Absent students</li>
                              <li>Open Gmail and click &ldquo;Compose&rdquo;</li>
                              <li>Paste (Ctrl+V) in the &ldquo;To&rdquo; field - all emails will be added automatically</li>
                              <li>Write your message and send to all students at once!</li>
                            </ol>
                          </div>
                        </div>
                      </div>

                      {/* Student Lists */}
                      <div className="mt-6 grid grid-cols-1 md:grid-cols-2 gap-4">
                        {/* Present Students */}
                        <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
                          <div className="flex items-center justify-between mb-3">
                            <div>
                              <h4 className="text-green-300 text-sm font-semibold">
                                Present Students ({attendanceStats.presentStudents.length})
                              </h4>
                              <p className="text-xs text-green-400/70 mt-1">
                                📧 Ready for Gmail &ldquo;To&rdquo; field
                              </p>
                            </div>
                            <button
                              onClick={copyPresentEmails}
                              disabled={attendanceStats.presentStudents.length === 0}
                              className="flex items-center gap-1 px-3 py-2 bg-green-600 hover:bg-green-700 text-white text-xs font-medium rounded transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                              title="Copy emails formatted for Gmail"
                            >
                              {copiedPresentEmails ? (
                                <>
                                  <Check className="w-3 h-3" />
                                  Copied for Gmail!
                                </>
                              ) : (
                                <>
                                  <Copy className="w-3 h-3" />
                                  Copy for Gmail
                                </>
                              )}
                            </button>
                          </div>
                          <div className="max-h-48 overflow-y-auto space-y-2">
                            {attendanceStats.presentStudents.length > 0 ? (
                              attendanceStats.presentStudents.map((student, index) => (
                                <div key={index} className="bg-green-700/20 rounded px-3 py-2">
                                  <div className="text-sm text-white font-medium">{student.name}</div>
                                  <div className="text-xs text-green-300 font-mono">{student.email}</div>
                                </div>
                              ))
                            ) : (
                              <div className="text-center text-green-400 text-sm py-4">
                                No students present
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Absent Students */}
                        <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
                          <div className="flex items-center justify-between mb-3">
                            <div>
                              <h4 className="text-red-300 text-sm font-semibold">
                                Absent Students ({attendanceStats.absentStudents.length})
                              </h4>
                              <p className="text-xs text-red-400/70 mt-1">
                                📧 Ready for Gmail &ldquo;To&rdquo; field
                              </p>
                            </div>
                            <button
                              onClick={copyAbsentEmails}
                              disabled={attendanceStats.absentStudents.length === 0}
                              className="flex items-center gap-1 px-3 py-2 bg-red-600 hover:bg-red-700 text-white text-xs font-medium rounded transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                              title="Copy emails formatted for Gmail"
                            >
                              {copiedAbsentEmails ? (
                                <>
                                  <Check className="w-3 h-3" />
                                  Copied for Gmail!
                                </>
                              ) : (
                                <>
                                  <Copy className="w-3 h-3" />
                                  Copy for Gmail
                                </>
                              )}
                            </button>
                          </div>
                          <div className="max-h-48 overflow-y-auto space-y-2">
                            {attendanceStats.absentStudents.length > 0 ? (
                              attendanceStats.absentStudents.map((student, index) => (
                                <div key={index} className="bg-red-700/20 rounded px-3 py-2">
                                  <div className="text-sm text-white font-medium">{student.name}</div>
                                  <div className="text-xs text-red-300 font-mono">{student.email}</div>
                                </div>
                              ))
                            ) : (
                              <div className="text-center text-red-400 text-sm py-4">
                                No students absent
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                // Regular Student Selection for other colleges
                <div className="space-y-4">
                  <div className="flex justify-between items-center">
                    <span className="text-sm text-gray-300">
                      {selectedStudents.size} of {studentData.length} students selected
                    </span>
                    <div className="flex gap-2">
                      <button 
                        onClick={() => setSelectedStudents(new Set(studentData.map((_, i) => i)))}
                        className="text-xs px-3 py-1 bg-blue-600 text-white rounded hover:bg-blue-700"
                      >
                        Select All
                      </button>
                      <button 
                        onClick={() => setSelectedStudents(new Set())}
                        className="text-xs px-3 py-1 bg-gray-600 text-white rounded hover:bg-gray-700"
                      >
                        Clear All
                      </button>
                    </div>
                  </div>
                  <div className="max-h-64 overflow-y-auto space-y-2">
                    {studentData.map((student, index) => (
                      <div 
                        key={index}
                        className="flex items-center justify-between p-3 bg-gray-700/50 rounded-lg hover:bg-gray-700/70 transition-colors"
                      >
                        <div className="flex items-center gap-3">
                          <input
                            type="checkbox"
                            checked={selectedStudents.has(index)}
                            onChange={() => toggleStudentSelection(index)}
                            className="w-4 h-4 text-blue-600 bg-gray-700 border-gray-600 rounded focus:ring-blue-500"
                          />
                          <div>
                            <div className="text-sm font-medium text-white">{student.name}</div>
                            <div className="text-xs text-gray-400">Roll: {student.rollNumber}</div>
                          </div>
                        </div>
                        <div className="text-right">
                          <div className="text-sm text-white">Rank: {student.rank}</div>
                          <div className="text-xs text-gray-400">{student.percentage}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                  <div className="flex justify-end pt-4">
                    <button 
                      onClick={generateEmails}
                      disabled={selectedStudents.size === 0}
                      className="px-6 py-2 bg-blue-600 text-white font-semibold rounded-md hover:bg-blue-700 transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                    >
                      Generate Emails ({selectedStudents.size} selected)
                    </button>
                  </div>
                </div>
              )
            ) : (
              <div className="text-center py-10">
                <Book className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                <p className="text-gray-500">
                  {!selectedCollege ? 'Select a college and upload a file...' : 'Upload a file to select students...'}
                </p>
              </div>
            )}
            </section>
          </div>

          {/* Row 2: Email Preview and Generated Emails */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            {selectedCollege && selectedStudents.size > 0 && (
              <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
              <div className="flex items-center gap-3 mb-4">
                <Mail className="w-5 h-5 text-gray-400" />
                <h2 className="text-lg font-semibold text-white">Email Preview</h2>
              </div>
              <div className="bg-gray-700/50 rounded-lg p-4">
                <h3 className="text-sm font-medium text-white mb-3">Email Template</h3>
                <div 
                  className="text-sm text-gray-300 whitespace-pre-wrap leading-relaxed [&_b]:font-bold [&_b]:text-white"
                  dangerouslySetInnerHTML={{
                    __html: `Dear {Name},

Greetings of the day!

We are pleased to share with you the results of your recent Daily Assessments till 18th August 2025 (15 sessions). <b>Your average percentage is {Percentage}</b>, reflecting your consistent effort and commitment.

We are also delighted to inform you that you have secured <b>Rank {Rank}</b> among your peers. Congratulations on this achievement!

Keep up the good work and continue striving for excellence.
With steady focus and dedication, you are sure to reach even greater milestones.

Wishing you all the best for your upcoming assessments.

If you have any questions feel free to reach out to ncetsupport@myanatomy.in

Warm regards,`.replace(/\n/g, '<br>')
                  }}
                />
              </div>
            </section>
            )}

            {/* Box 4: Generated Emails Preview */}
          {generatedEmails.length > 0 && (
            <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-3">
                  <Mail className="w-5 h-5 text-gray-400" />
                  <h2 className="text-lg font-semibold text-white">Generated Emails</h2>
                  <span className="bg-blue-600 text-white text-xs px-2 py-1 rounded-full">
                    {generatedEmails.length} emails
                  </span>
                  <span className="bg-green-600 text-white text-xs px-2 py-1 rounded-full">
                    {emailsSent.size} sent
                  </span>
                </div>
                <div className="flex gap-2">
                  <button 
                    onClick={markAllEmailsSent}
                    className="text-xs px-3 py-1 bg-green-600 text-white rounded hover:bg-green-700 transition-colors"
                  >
                    Mark All Sent
                  </button>
                  <button 
                    onClick={clearAllEmailsSent}
                    className="text-xs px-3 py-1 bg-gray-600 text-white rounded hover:bg-gray-700 transition-colors"
                  >
                    Clear All
                  </button>
                </div>
              </div>
              <div className="mb-4 p-3 bg-blue-600/20 border border-blue-500/30 rounded-lg">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-sm font-medium text-white">Subject Line (Common for all emails):</div>
                  <button 
                    onClick={copySubjectLine}
                    className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors"
                  >
                    {copiedSubject ? (
                      <>
                        <Check className="w-3 h-3" />
                        Copied!
                      </>
                    ) : (
                      <>
                        <Copy className="w-3 h-3" />
                        Copy
                      </>
                    )}
                  </button>
                </div>
                <div className="text-sm text-blue-300 font-medium">Congratulations on Your Daily Assessment Performance</div>
              </div>
              <div className="mb-4 p-3 bg-green-600/20 border border-green-500/30 rounded-lg">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-xs font-medium text-white">📧 Email Progress:</div>
                  <div className="text-xs text-green-300">
                    {emailsSent.size} of {generatedEmails.length} emails sent ({Math.round((emailsSent.size / generatedEmails.length) * 100)}%)
                  </div>
                </div>
                <div className="w-full bg-gray-700 rounded-full h-2">
                  <div 
                    className="bg-green-600 h-2 rounded-full transition-all duration-300"
                    style={{ width: `${(emailsSent.size / generatedEmails.length) * 100}%` }}
                  ></div>
                </div>
              </div>
              <div className="space-y-4">
                <div className="max-h-96 overflow-y-auto space-y-4">
                  {generatedEmails.map((emailData, index) => {
                    const selectedIndices = Array.from(selectedStudents);
                    const originalIndex = selectedIndices[index];
                    const student = studentData[originalIndex];
                    return (
                      <div key={index} className={`rounded-lg p-4 transition-all ${emailsSent.has(originalIndex) ? 'bg-green-700/30 border border-green-500/30' : 'bg-gray-700/50'}`}>
                        <div className="flex items-center justify-between mb-3">
                          <div className="flex items-center gap-3">
                            <input
                              type="checkbox"
                              checked={emailsSent.has(originalIndex)}
                              onChange={() => toggleEmailSent(originalIndex)}
                              className="w-4 h-4 text-green-600 bg-gray-700 border-gray-600 rounded focus:ring-green-500"
                            />
                            <h3 className={`text-sm font-medium ${emailsSent.has(originalIndex) ? 'text-green-300 line-through' : 'text-white'}`}>
                              Email {index + 1} - {student?.name} (Rank: {student?.rank})
                              {emailsSent.has(originalIndex) && <span className="ml-2 text-xs bg-green-600 text-white px-2 py-0.5 rounded-full">Sent</span>}
                            </h3>
                          </div>
                          <div className="flex gap-2">
                            <button 
                              onClick={() => copyToClipboard(emailData, index)}
                              className="flex items-center gap-2 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded-md transition-colors"
                            >
                              {copiedEmailIndex === index ? (
                                <>
                                  <Check className="w-3 h-3" />
                                  Copied!
                                </>
                              ) : (
                                <>
                                  <Copy className="w-3 h-3" />
                                  Copy Email
                                </>
                              )}
                            </button>
                            <button 
                              onClick={() => toggleEmailSent(originalIndex)}
                              className={`flex items-center gap-2 px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
                                emailsSent.has(originalIndex)
                                  ? 'bg-green-600 hover:bg-green-700 text-white'
                                  : 'bg-gray-600 hover:bg-gray-700 text-white'
                              }`}
                            >
                              {emailsSent.has(originalIndex) ? (
                                <>
                                  <Check className="w-3 h-3" />
                                  Sent
                                </>
                              ) : (
                                <>
                                  <Mail className="w-3 h-3" />
                                  Mark Sent
                                </>
                              )}
                            </button>
                          </div>
                        </div>
                        <div 
                          className="text-xs text-gray-300 whitespace-pre-wrap bg-gray-800 rounded p-3 overflow-x-auto leading-relaxed font-mono [&_b]:font-bold [&_b]:text-white"
                          dangerouslySetInnerHTML={{ __html: emailData.displayVersion.replace(/\n/g, '<br>') }}
                        />
                      </div>
                    );
                  })}
                </div>
                <div className="flex justify-between items-center pt-4 border-t border-gray-600">
                  <div className="text-sm text-gray-400">
                    {generatedEmails.length} emails ready to send
                  </div>
                  <div className="flex gap-3">
                    <button 
                      onClick={() => setGeneratedEmails([])}
                      className="px-4 py-2 bg-gray-600 text-white rounded-md hover:bg-gray-700 transition-colors"
                    >
                      Clear All
                    </button>
                    <button 
                      onClick={copyAllEmails}
                      className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
                    >
                      {copiedEmailIndex === -1 ? (
                        <>
                          <Check className="w-4 h-4" />
                          Copied All!
                        </>
                      ) : (
                        <>
                          <Copy className="w-4 h-4" />
                          Copy All Emails
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            </section>
          )}
          </div>

          {/* Row 4: Email Template Generator - Full Width */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-3">
                <Mail className="w-5 h-5 text-gray-400" />
                <h2 className="text-lg font-semibold text-white">Generated Email Template</h2>
                <span className="bg-orange-600 text-white text-xs px-2 py-1 rounded-full">
                  Attendance Summary
                </span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={generateEmailTemplate}
                  className="flex items-center gap-2 px-3 py-1.5 bg-orange-600 hover:bg-orange-700 text-white text-sm rounded-md transition-colors"
                >
                  Generate Template
                </button>
                <button
                  onClick={copyEmailTemplate}
                  disabled={!emailTemplate.generatedContent}
                  className="flex items-center gap-2 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-sm rounded-md transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                >
                  {copiedEmailTemplate ? (
                    <>
                      <Check className="w-4 h-4" />
                      Copied!
                    </>
                  ) : (
                    <>
                      <Copy className="w-4 h-4" />
                      Copy Email
                    </>
                  )}
                </button>
              </div>
            </div>
            
            <p className="text-sm text-gray-400 mb-6">
              Generate professional attendance summary emails for administrators with customizable batch data and training dates.
            </p>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              {/* Settings Panel */}
              <div className="lg:col-span-1 space-y-4">
                <div className="bg-gray-700/50 rounded-lg p-4">
                  <h3 className="text-sm font-semibold text-white mb-4">Email Settings</h3>
                  
                  {/* Training Date */}
                  <div className="space-y-2">
                    <label className="text-xs text-gray-300 font-medium">Training Date:</label>
                    <input
                      type="text"
                      value={emailTemplate.trainingDate}
                      onChange={(e) => setEmailTemplate(prev => ({ ...prev, trainingDate: e.target.value }))}
                      placeholder="DD/MM/YYYY"
                      className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-orange-500 focus:border-orange-500 select-text"
                      style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                    />
                  </div>
                  
                  {/* Google Sheets Link */}
                  <div className="space-y-2">
                    <label className="text-xs text-gray-300 font-medium">Google Sheets Link:</label>
                    <input
                      type="text"
                      value={emailTemplate.sheetsLink}
                      onChange={(e) => setEmailTemplate(prev => ({ ...prev, sheetsLink: e.target.value }))}
                      placeholder="https://docs.google.com/spreadsheets/d/..."
                      className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-orange-500 focus:border-orange-500 select-text"
                      style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                    />
                  </div>
                </div>

                {/* Batch Data Display */}
                <div className="bg-gray-700/50 rounded-lg p-4">
                  <h3 className="text-sm font-semibold text-white mb-4">Batch Attendance Data</h3>
                  {selectedDate ? (
                    <div className="space-y-3">
                      {emailTemplate.batches.length > 0 ? (
                        <>
                          <div className="flex items-center gap-2 mb-3">
                            <div className="w-2 h-2 bg-green-500 rounded-full"></div>
                            <span className="text-xs text-green-300 font-medium">
                              Using data from {allSheetsAttendanceData.size} selected sheets for {selectedDate}
                            </span>
                          </div>
                          {emailTemplate.batches.map((batch) => (
                            <div key={batch.id} className="bg-gray-600/50 rounded p-3">
                              <div className="text-sm font-medium text-white mb-2">{batch.name}</div>
                              <div className="grid grid-cols-3 gap-4">
                                <div className="text-center">
                                  <div className="text-xs text-gray-400">Total</div>
                                  <div className="text-lg font-bold text-blue-300">{batch.total}</div>
                                </div>
                                <div className="text-center">
                                  <div className="text-xs text-gray-400">Present</div>
                                  <div className="text-lg font-bold text-green-300">{batch.present}</div>
                                </div>
                                <div className="text-center">
                                  <div className="text-xs text-gray-400">Absent</div>
                                  <div className="text-lg font-bold text-red-300">{batch.absent}</div>
                                </div>
                              </div>
                            </div>
                          ))}
                        </>
                      ) : (
                        <div className="bg-orange-600/20 border border-orange-500/30 rounded p-3 text-center">
                          <div className="text-orange-300 text-sm font-medium mb-1">No Batch Data Available</div>
                          <div className="text-orange-200/80 text-xs">
                            Please select sheets to process and click &ldquo;Generate Template&rdquo; to load attendance data for {selectedDate}.
                          </div>
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className="bg-orange-600/20 border border-orange-500/30 rounded p-4 text-center">
                      <div className="text-orange-300 text-sm font-medium mb-1">No Date Selected</div>
                      <div className="text-orange-200/80 text-xs">
                        Please select NIET college, upload attendance sheet, choose sheets to process, and select a date to generate email template.
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* Generated Email Preview */}
              <div className="lg:col-span-2">
                <div className="bg-gray-700/50 rounded-lg p-4">
                  <h3 className="text-sm font-semibold text-white mb-4">Email Preview</h3>
                  
                  {/* Subject Line */}
                  {emailTemplateSubject && (
                    <div className="bg-blue-600/10 border border-blue-500/30 rounded-lg p-3 mb-4">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <Mail className="w-4 h-4 text-blue-400" />
                          <span className="text-xs font-semibold text-blue-300">Subject Line</span>
                        </div>
                        <button
                          onClick={copyEmailTemplateSubject}
                          className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors"
                        >
                          {copiedEmailTemplateSubject ? (
                            <>
                              <Check className="w-3 h-3" />
                              Copied!
                            </>
                          ) : (
                            <>
                              <Copy className="w-3 h-3" />
                              Copy Subject
                            </>
                          )}
                        </button>
                      </div>
                      <div className="mt-2 text-sm text-blue-100 font-medium select-text" style={{ userSelect: 'text', WebkitUserSelect: 'text' }}>
                        {emailTemplateSubject}
                      </div>
                    </div>
                  )}
                  
                  {emailTemplate.generatedContent ? (
                    <div className="bg-gray-800 rounded-lg p-4 border border-gray-600">
                      <div 
                        className="text-sm text-gray-200 leading-relaxed select-text"
                        style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                        dangerouslySetInnerHTML={{ __html: emailTemplate.generatedContent }}
                      />
                    </div>
                  ) : (
                    <div className="bg-gray-800 rounded-lg p-8 border-2 border-dashed border-gray-600 text-center">
                      <Mail className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                      <p className="text-gray-500 font-medium mb-2">Email template not generated yet</p>
                      <p className="text-xs text-gray-600">
                        Click &ldquo;Generate Template&rdquo; to create your attendance summary email
                      </p>
                    </div>
                  )}
                </div>
                
                {/* Help Section */}
                <div className="bg-orange-600/10 border border-orange-500/30 rounded-lg p-4 mt-4">
                  <div className="flex items-start gap-3">
                    <Mail className="w-5 h-5 text-orange-400 mt-0.5 flex-shrink-0" />
                    <div>
                      <h4 className="text-orange-300 text-sm font-semibold mb-2">How to Use Email Template</h4>
                      <ul className="text-xs text-orange-200/80 space-y-1 list-disc list-inside">
                        <li>Customize training date and batch data in the settings panel</li>
                        <li>Update Google Sheets link to match your attendance document</li>
                        <li>Click &ldquo;Generate Template&rdquo; to create the email content with formatting</li>
                        <li>Use &ldquo;Copy Email&rdquo; to copy HTML-formatted content to your clipboard</li>
                        <li><strong>Gmail users:</strong> Paste directly into Gmail compose window - formatting will be preserved</li>
                        <li><strong>Bold text:</strong> Attendance data, key phrases, and links will appear bold</li>
                        <li><strong>Indentation:</strong> Attendance statistics will be properly indented</li>
                        <li>Add your signature details before sending</li>
                      </ul>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </section>

                    {/* Row 3: Intern Report Section - Full Width */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-3">
                <FileText className="w-5 h-5 text-gray-400" />
                <h2 className="text-lg font-semibold text-white">Intern Report</h2>
                <span className="bg-purple-600 text-white text-xs px-2 py-1 rounded-full">
                  Text Format
                </span>
              </div>
              <button
                onClick={() => setInternReportExpanded(!internReportExpanded)}
                className="flex items-center gap-2 px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white text-sm rounded-md transition-colors"
              >
                {internReportExpanded ? 'Collapse' : 'Expand'}
                <ChevronDown className={`w-4 h-4 transition-transform ${internReportExpanded ? 'rotate-180' : ''}`} />
              </button>
            </div>
            
            <p className="text-sm text-gray-400 mb-4">
              Paste your intern report in text format below. The system will process and organize the content automatically.
            </p>

            <div className="space-y-4">
              {/* Text Input Area */}
              <div className="w-full">
                <label htmlFor="intern-report-input" className="text-sm text-gray-300 font-medium mb-2 block">
                  Intern Report Content:
                </label>
                <textarea
                  id="intern-report-input"
                  value={internReport}
                  onChange={(e) => {
                    setInternReport(e.target.value);
                    processInternReport(e.target.value);
                  }}
                  placeholder="Paste your intern report content here...&#10;&#10;You can include:&#10;- Student names and details&#10;- Internship progress&#10;- Performance evaluations&#10;- Any other relevant information"
                  className="w-full h-48 bg-gray-700 border border-gray-600 rounded-lg px-4 py-3 text-white placeholder-gray-400 focus:ring-2 focus:ring-purple-500 focus:border-purple-500 resize-y select-text"
                  style={{ userSelect: 'text', WebkitUserSelect: 'text', minHeight: '12rem' }}
                />
                <div className="flex justify-between items-center mt-2">
                  <span className="text-xs text-gray-500">
                    {internReport.length} characters, {internReport.split('\n').filter(line => line.trim()).length} lines
                  </span>
                  {internReport.trim() && (
                    <button
                      onClick={() => {
                        setInternReport('');
                        setProcessedInternData([]);
                      }}
                      className="text-xs px-2 py-1 bg-gray-600 hover:bg-gray-700 text-white rounded transition-colors"
                    >
                      Clear
                    </button>
                  )}
                </div>
              </div>

              {/* Processed Data Display */}
              {processedInternData.length > 0 && (
                <div className={`transition-all duration-300 ${internReportExpanded ? 'block' : 'hidden'}`}>
                  <div className="bg-gray-700/50 rounded-lg p-4">
                    <h3 className="text-sm font-semibold text-white mb-3">
                      Processed Report Data ({processedInternData.length} entries)
                    </h3>
                    <div className="max-h-64 overflow-y-auto space-y-2">
                      {processedInternData.map((item) => (
                        <div key={item.id} className="bg-gray-800/50 rounded px-3 py-2">
                          <div className="flex items-start justify-between">
                            <div className="flex-1">
                              <span className="text-xs text-purple-400">#{item.id}</span>
                              <p className="text-sm text-gray-200 mt-1">{item.content}</p>
                            </div>
                            <span className="text-xs px-2 py-1 bg-purple-600/20 text-purple-300 rounded ml-3">
                              {item.type}
                            </span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* Help Section */}
              <div className="bg-purple-600/10 border border-purple-500/30 rounded-lg p-4">
                <div className="flex items-start gap-3">
                  <FileText className="w-5 h-5 text-purple-400 mt-0.5 flex-shrink-0" />
                  <div>
                    <h4 className="text-purple-300 text-sm font-semibold mb-2">How to Use Intern Report</h4>
                    <ul className="text-xs text-purple-200/80 space-y-1 list-disc list-inside">
                      <li>Copy your intern report from any document or email</li>
                      <li>Paste the content in the text area above</li>
                      <li>The system will automatically process and organize the data</li>
                      <li>Use &ldquo;Expand&rdquo; to view the processed entries</li>
                      <li>Each line of meaningful content becomes a separate entry</li>
                    </ul>
                  </div>
                </div>
              </div>
            </div>
          </section>

          {/* Row 5: Student Email Templates - Full Width */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-3">
                <Mail className="w-5 h-5 text-gray-400" />
                <h2 className="text-lg font-semibold text-white">Student Email Templates</h2>
                <span className="bg-purple-600 text-white text-xs px-2 py-1 rounded-full">
                  For Students
                </span>
              </div>
            </div>
            
            <p className="text-sm text-gray-400 mb-4">
              Generate personalized email templates for students based on their attendance status. Tables are auto-generated from intern report data.
            </p>

            {/* Batch Selection for Email Templates */}
            <div className="mb-6 p-4 bg-purple-600/10 border border-purple-500/30 rounded-lg">
              <div className="flex items-center gap-3 mb-3">
                <Users className="w-4 h-4 text-purple-400" />
                <h4 className="text-sm font-semibold text-purple-300">Select Batch for Email Template</h4>
              </div>
              <p className="text-xs text-purple-200/70 mb-3">
                Choose which batch topic to include in the email template table (will show only 1 row)
              </p>
              <select
                value={selectedBatchForEmail}
                onChange={(e) => setSelectedBatchForEmail(e.target.value)}
                className="w-full px-3 py-2 bg-gray-700 border border-gray-600 rounded-md text-white text-sm focus:ring-2 focus:ring-purple-500 focus:border-transparent"
              >
                <option value="">Select a batch...</option>
                {Array.from(selectedSheetsForProcessing).map((sheetName: string) => (
                  <option key={sheetName} value={sheetName}>
                    {sheetName}
                  </option>
                ))}
              </select>
              {selectedBatchForEmail === '' && (
                <p className="text-xs text-yellow-400 mt-2 flex items-center gap-1">
                  <AlertCircle className="w-3 h-3" />
                  Please select a batch to generate email templates with table data
                </p>
              )}
            </div>

            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              {/* Absent Students Email */}
              <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="text-lg font-semibold text-red-300 mb-2">Email for Absent Students</h3>
                    <p className="text-xs text-red-400/70">For students who missed the training session</p>
                  </div>
                  <div className="flex gap-2">
                    <button
                      onClick={generateAbsentStudentEmail}
                      className="flex items-center gap-2 px-3 py-1.5 bg-red-600 hover:bg-red-700 text-white text-xs font-medium rounded-md transition-colors"
                    >
                      Generate Email
                    </button>
                    <button
                      onClick={copyAbsentStudentEmail}
                      disabled={!absentStudentEmailContent}
                      className="flex items-center gap-2 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded-md transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                    >
                      {copiedAbsentStudentEmail ? (
                        <>
                          <Check className="w-3 h-3" />
                          Copied!
                        </>
                      ) : (
                        <>
                          <Copy className="w-3 h-3" />
                          Copy Email
                        </>
                      )}
                    </button>
                  </div>
                </div>
                
                {/* Subject Line */}
                {absentStudentEmailSubject && (
                  <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-3 mb-4">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <Mail className="w-4 h-4 text-red-400" />
                        <span className="text-xs font-semibold text-red-300">Subject Line</span>
                      </div>
                      <button
                        onClick={copyAbsentStudentEmailSubject}
                        className="flex items-center gap-1 px-2 py-1 bg-red-600 hover:bg-red-700 text-white text-xs font-medium rounded transition-colors"
                      >
                        {copiedAbsentStudentSubject ? (
                          <>
                            <Check className="w-3 h-3" />
                            Copied!
                          </>
                        ) : (
                          <>
                            <Copy className="w-3 h-3" />
                            Copy Subject
                          </>
                        )}
                      </button>
                    </div>
                    <div className="mt-2 text-sm text-red-100 font-medium select-text" style={{ userSelect: 'text', WebkitUserSelect: 'text' }}>
                      {absentStudentEmailSubject}
                    </div>
                  </div>
                )}
                
                {absentStudentEmailContent ? (
                  <div className="bg-gray-800 rounded-lg p-4 border border-gray-600 max-h-96 overflow-y-auto">
                    <div 
                      className="text-sm text-gray-200 leading-relaxed select-text"
                      style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                      dangerouslySetInnerHTML={{ __html: absentStudentEmailContent }}
                    />
                  </div>
                ) : (
                  <div className="bg-gray-800 rounded-lg p-8 border-2 border-dashed border-gray-600 text-center">
                    <Mail className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                    <p className="text-gray-500 font-medium mb-2">Absent student email not generated yet</p>
                    <p className="text-xs text-gray-600">
                      Click &ldquo;Generate Email&rdquo; to create the absent student email template
                    </p>
                  </div>
                )}
              </div>

              {/* Present Students Email */}
              <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="text-lg font-semibold text-green-300 mb-2">Email for Present Students</h3>
                    <p className="text-xs text-green-400/70">For students who attended the training session</p>
                  </div>
                  <div className="flex gap-2">
                    <button
                      onClick={generatePresentStudentEmail}
                      className="flex items-center gap-2 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white text-xs font-medium rounded-md transition-colors"
                    >
                      Generate Email
                    </button>
                    <button
                      onClick={copyPresentStudentEmail}
                      disabled={!presentStudentEmailContent}
                      className="flex items-center gap-2 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded-md transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                    >
                      {copiedPresentStudentEmail ? (
                        <>
                          <Check className="w-3 h-3" />
                          Copied!
                        </>
                      ) : (
                        <>
                          <Copy className="w-3 h-3" />
                          Copy Email
                        </>
                      )}
                    </button>
                  </div>
                </div>
                
                {/* Subject Line */}
                {presentStudentEmailSubject && (
                  <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-3 mb-4">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <Mail className="w-4 h-4 text-green-400" />
                        <span className="text-xs font-semibold text-green-300">Subject Line</span>
                      </div>
                      <button
                        onClick={copyPresentStudentEmailSubject}
                        className="flex items-center gap-1 px-2 py-1 bg-green-600 hover:bg-green-700 text-white text-xs font-medium rounded transition-colors"
                      >
                        {copiedPresentStudentSubject ? (
                          <>
                            <Check className="w-3 h-3" />
                            Copied!
                          </>
                        ) : (
                          <>
                            <Copy className="w-3 h-3" />
                            Copy Subject
                          </>
                        )}
                      </button>
                    </div>
                    <div className="mt-2 text-sm text-green-100 font-medium select-text" style={{ userSelect: 'text', WebkitUserSelect: 'text' }}>
                      {presentStudentEmailSubject}
                    </div>
                  </div>
                )}
                
                {presentStudentEmailContent ? (
                  <div className="bg-gray-800 rounded-lg p-4 border border-gray-600 max-h-96 overflow-y-auto">
                    <div 
                      className="text-sm text-gray-200 leading-relaxed select-text"
                      style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                      dangerouslySetInnerHTML={{ __html: presentStudentEmailContent }}
                    />
                  </div>
                ) : (
                  <div className="bg-gray-800 rounded-lg p-8 border-2 border-dashed border-gray-600 text-center">
                    <Mail className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                    <p className="text-gray-500 font-medium mb-2">Present student email not generated yet</p>
                    <p className="text-xs text-gray-600">
                      Click &ldquo;Generate Email&rdquo; to create the present student email template
                    </p>
                  </div>
                )}
              </div>
            </div>

            {/* Help Section */}
            <div className="mt-6 bg-purple-600/10 border border-purple-500/30 rounded-lg p-4">
              <div className="flex items-start gap-3">
                <Mail className="w-5 h-5 text-purple-400 mt-0.5 flex-shrink-0" />
                <div>
                  <h4 className="text-purple-300 text-sm font-semibold mb-2">How to Use Student Email Templates</h4>
                  <ul className="text-xs text-purple-200/80 space-y-1 list-disc list-inside">
                    <li><strong>Absent Students:</strong> Generates email with session recap and attendance warning</li>
                    <li><strong>Present Students:</strong> Generates congratulatory email with session summary</li>
                    <li><strong>Dynamic Tables:</strong> Topics use selected sheet names, descriptions from intern report</li>
                    <li><strong>HTML Formatting:</strong> Tables and bold text will be preserved when pasted in Gmail</li>
                    <li><strong>Date Integration:</strong> Uses the same date selected in attendance analysis</li>
                    <li><strong>Mass Emailing:</strong> Copy template and send to respective student groups</li>
                  </ul>
                </div>
              </div>
            </div>
          </section>

          {/* Row 6: Collective Students - Full Width */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-3">
                <Users className="w-5 h-5 text-gray-400" />
                <h2 className="text-lg font-semibold text-white">Collective Students</h2>
                <span className="bg-cyan-600 text-white text-xs px-2 py-1 rounded-full">
                  All Batches
                </span>
              </div>
            </div>
            
            <p className="text-sm text-gray-400 mb-6">
              View and copy email addresses of all students (present/absent) across all selected batches for the chosen date.
            </p>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* All Present Students */}
              <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="text-lg font-semibold text-green-300 mb-2">All Present Students</h3>
                    <p className="text-xs text-green-400/70">Students who attended from all batches</p>
                  </div>
                  <button
                    onClick={copyAllPresentStudentEmails}
                    disabled={getAllPresentStudents().length === 0}
                    className="flex items-center gap-2 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white text-xs font-medium rounded-md transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                  >
                    {copiedPresentEmails ? (
                      <>
                        <Check className="w-3 h-3" />
                        Copied!
                      </>
                    ) : (
                      <>
                        <Copy className="w-3 h-3" />
                        Copy All Present Emails
                      </>
                    )}
                  </button>
                </div>
                
                <div className="bg-gray-800 rounded-lg p-4 border border-gray-600 max-h-64 overflow-y-auto">
                  {getAllPresentStudents().length > 0 ? (
                    <div className="space-y-2">
                      <div className="text-sm font-semibold text-green-300 mb-3">
                        Total Present: {getAllPresentStudents().length} students
                      </div>
                      {getAllPresentStudents().map((student, index) => (
                        <div key={index} className="flex justify-between items-center text-xs text-gray-300 bg-gray-700/50 p-2 rounded">
                          <span className="font-medium">{student.name}</span>
                          <span className="text-green-400">{student.email}</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="text-center text-gray-500 py-8">
                      <Users className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                      <p className="text-sm font-medium mb-2">No present students found</p>
                      <p className="text-xs">Process attendance data to see present students</p>
                    </div>
                  )}
                </div>
              </div>

              {/* All Absent Students */}
              <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="text-lg font-semibold text-red-300 mb-2">All Absent Students</h3>
                    <p className="text-xs text-red-400/70">Students who missed from all batches</p>
                  </div>
                  <button
                    onClick={copyAllAbsentStudentEmails}
                    disabled={getAllAbsentStudents().length === 0}
                    className="flex items-center gap-2 px-3 py-1.5 bg-red-600 hover:bg-red-700 text-white text-xs font-medium rounded-md transition-colors disabled:bg-gray-600 disabled:cursor-not-allowed"
                  >
                    {copiedAbsentEmails ? (
                      <>
                        <Check className="w-3 h-3" />
                        Copied!
                      </>
                    ) : (
                      <>
                        <Copy className="w-3 h-3" />
                        Copy All Absent Emails
                      </>
                    )}
                  </button>
                </div>
                
                <div className="bg-gray-800 rounded-lg p-4 border border-gray-600 max-h-64 overflow-y-auto">
                  {getAllAbsentStudents().length > 0 ? (
                    <div className="space-y-2">
                      <div className="text-sm font-semibold text-red-300 mb-3">
                        Total Absent: {getAllAbsentStudents().length} students
                      </div>
                      {getAllAbsentStudents().map((student, index) => (
                        <div key={index} className="flex justify-between items-center text-xs text-gray-300 bg-gray-700/50 p-2 rounded">
                          <span className="font-medium">{student.name}</span>
                          <span className="text-red-400">{student.email}</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="text-center text-gray-500 py-8">
                      <Users className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                      <p className="text-sm font-medium mb-2">No absent students found</p>
                      <p className="text-xs">Process attendance data to see absent students</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Summary Information */}
            <div className="mt-6 bg-cyan-600/10 border border-cyan-500/30 rounded-lg p-4">
              <div className="flex items-start gap-3">
                <Users className="w-5 h-5 text-cyan-400 mt-0.5 flex-shrink-0" />
                <div>
                  <h4 className="text-cyan-300 text-sm font-semibold mb-2">Collective Summary</h4>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-xs">
                    <div className="text-center">
                      <div className="text-2xl font-bold text-green-400">{getAllPresentStudents().length}</div>
                      <div className="text-gray-400">Total Present</div>
                    </div>
                    <div className="text-center">
                      <div className="text-2xl font-bold text-red-400">{getAllAbsentStudents().length}</div>
                      <div className="text-gray-400">Total Absent</div>
                    </div>
                    <div className="text-center">
                      <div className="text-2xl font-bold text-cyan-400">{getAllPresentStudents().length + getAllAbsentStudents().length}</div>
                      <div className="text-gray-400">Total Students</div>
                    </div>
                    <div className="text-center">
                      <div className="text-2xl font-bold text-purple-400">{selectedSheetsForProcessing.size}</div>
                      <div className="text-gray-400">Active Batches</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </section>

        </main>
      </div>
    </div>
  );
}