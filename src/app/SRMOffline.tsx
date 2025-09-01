'use client';

import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { FileText, Upload, Book, Copy, Check, Mail, ClipboardList } from 'lucide-react';

interface SRMOfflineStudent {
  serialNo: string;
  name: string;
  email: string;
  regnNumber: string;
  contactNumber: string;
  attendance: { [key: string]: number }; // date -> 0/1
}

interface SRMOfflineAttendanceStats {
  date: string;
  totalStudents: number;
  present: number;
  absent: number;
  presentPercentage: number;
  absentPercentage: number;
  presentStudents: Array<{ name: string; email: string; regnNumber: string }>;
  absentStudents: Array<{ name: string; email: string; regnNumber: string }>;
}

interface EmailTemplate {
  trainingDate: string;
  batches: string[];
  sheetsLink: string;
  to: string;
  generatedContent: string;
}

interface SRMOfflineProps {
  isVisible: boolean;
}

export default function SRMOffline({ isVisible }: SRMOfflineProps) {
  // File and data states
  const [attendanceFile, setAttendanceFile] = useState<File | null>(null);
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [selectedSheetsForProcessing, setSelectedSheetsForProcessing] = useState<Set<string>>(new Set());
  const [allSheetsData, setAllSheetsData] = useState<Map<string, SRMOfflineStudent[]>>(new Map());
  const [allSheetsAttendanceData, setAllSheetsAttendanceData] = useState<Map<string, SRMOfflineAttendanceStats>>(new Map());
  const [attendanceDates, setAttendanceDates] = useState<Array<{ date: string; fullText: string }>>([]);
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [attendanceStats, setAttendanceStats] = useState<SRMOfflineAttendanceStats | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploadComplete, setIsUploadComplete] = useState(false);

  // Email states
  const [absentStudentEmailContent, setAbsentStudentEmailContent] = useState<string>('');
  const [presentStudentEmailContent, setPresentStudentEmailContent] = useState<string>('');
  const [absentStudentEmailSubject, setAbsentStudentEmailSubject] = useState<string>('');
  const [presentStudentEmailSubject, setPresentStudentEmailSubject] = useState<string>('');
  const [absentStudentEmailTo, setAbsentStudentEmailTo] = useState<string>('');
  const [presentStudentEmailTo, setPresentStudentEmailTo] = useState<string>('');
  const [absentStudentEmailCC, setAbsentStudentEmailCC] = useState<string>('');
  const [presentStudentEmailCC, setPresentStudentEmailCC] = useState<string>('');
  const [absentStudentEmailBCC, setAbsentStudentEmailBCC] = useState<string>('');
  const [presentStudentEmailBCC, setPresentStudentEmailBCC] = useState<string>('');

  // Copy states
  const [copiedAbsentStudentEmail, setCopiedAbsentStudentEmail] = useState<boolean>(false);
  const [copiedPresentStudentEmail, setCopiedPresentStudentEmail] = useState<boolean>(false);
  const [copiedAbsentEmails, setCopiedAbsentEmails] = useState<boolean>(false);
  const [copiedPresentEmails, setCopiedPresentEmails] = useState<boolean>(false);

  // Email template states
  const [emailTemplate, setEmailTemplate] = useState<EmailTemplate>({
    trainingDate: '',
    batches: [],
    sheetsLink: 'https://docs.google.com/spreadsheets/d/1SRM_OFFLINE_LINK/edit?usp=sharing',
    to: 'offline.coordinator@srm.ac.in, dean@srm.ac.in, hod@srm.ac.in, admin@srm.ac.in',
    generatedContent: ''
  });
  const [copiedEmailTemplate, setCopiedEmailTemplate] = useState<boolean>(false);
  const [copiedEmailTo, setCopiedEmailTo] = useState<boolean>(false);

  // Intern report states
  const [internReport, setInternReport] = useState<string>('');
  const [internReportExpanded, setInternReportExpanded] = useState<boolean>(false);

  const calculateAttendanceStats = useCallback((date: string): SRMOfflineAttendanceStats | null => {
    if (!selectedSheetsForProcessing.size || !allSheetsData.size) return null;
    
    const allStudents: SRMOfflineStudent[] = [];
    
    // Combine students from selected sheets
    selectedSheetsForProcessing.forEach(sheetName => {
      const sheetData = allSheetsData.get(sheetName);
      if (sheetData) {
        allStudents.push(...sheetData);
      }
    });
    
    if (allStudents.length === 0) return null;
    
    const presentStudents: Array<{ name: string; email: string; regnNumber: string }> = [];
    const absentStudents: Array<{ name: string; email: string; regnNumber: string }> = [];
    
    allStudents.forEach(student => {
      const attendanceValue = student.attendance[date];
      if (attendanceValue === 1) {
        presentStudents.push({
          name: student.name,
          email: student.email,
          regnNumber: student.regnNumber
        });
      } else if (attendanceValue === 0) {
        absentStudents.push({
          name: student.name,
          email: student.email,
          regnNumber: student.regnNumber
        });
      }
    });
    
    const totalStudents = presentStudents.length + absentStudents.length;
    const presentPercentage = totalStudents > 0 ? Math.round((presentStudents.length / totalStudents) * 100) : 0;
    const absentPercentage = totalStudents > 0 ? Math.round((absentStudents.length / totalStudents) * 100) : 0;
    
    return {
      date,
      totalStudents,
      present: presentStudents.length,
      absent: absentStudents.length,
      presentPercentage,
      absentPercentage,
      presentStudents,
      absentStudents
    };
  }, [selectedSheetsForProcessing, allSheetsData]);

  useEffect(() => {
    if (selectedDate && selectedSheetsForProcessing.size > 0) {
      const stats = calculateAttendanceStats(selectedDate);
      setAttendanceStats(stats);
      
      // Calculate stats for all sheets
      const allSheetStats = new Map<string, SRMOfflineAttendanceStats>();
      selectedSheetsForProcessing.forEach(sheetName => {
        const sheetData = allSheetsData.get(sheetName);
        if (sheetData) {
          const presentStudents = sheetData.filter(student => student.attendance[selectedDate] === 1)
            .map(student => ({ name: student.name, email: student.email, regnNumber: student.regnNumber }));
          const absentStudents = sheetData.filter(student => student.attendance[selectedDate] === 0)
            .map(student => ({ name: student.name, email: student.email, regnNumber: student.regnNumber }));
          
          const totalStudents = presentStudents.length + absentStudents.length;
          const presentPercentage = totalStudents > 0 ? Math.round((presentStudents.length / totalStudents) * 100) : 0;
          const absentPercentage = totalStudents > 0 ? Math.round((absentStudents.length / totalStudents) * 100) : 0;
          
          allSheetStats.set(sheetName, {
            date: selectedDate,
            totalStudents,
            present: presentStudents.length,
            absent: absentStudents.length,
            presentPercentage,
            absentPercentage,
            presentStudents,
            absentStudents
          });
        }
      });
      setAllSheetsAttendanceData(allSheetStats);
    }
  }, [selectedDate, selectedSheetsForProcessing, allSheetsData, calculateAttendanceStats]);

  if (!isVisible) return null;

  // Functions
  const loadSRMOfflineAttendanceSheet = () => {
    // This would load a pre-configured attendance sheet for SRM Offline
    console.log('Loading SRM Offline attendance sheet...');
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file && file.name.endsWith('.xlsx')) {
      setAttendanceFile(file);
      processAttendanceFile(file);
    }
  };

  const processAttendanceFile = (file: File) => {
    setIsProcessing(true);
    setIsUploadComplete(false);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        
        console.log('SRM Offline - Workbook sheet names:', workbook.SheetNames);
        
        // Get available sheets and filter to only include the SRM Offline batch names
        const allSheetNames = workbook.SheetNames;
        const validSheetNames = ['MS1', 'MS2', 'AI/ML-1', 'AI/ML-2', 'AI/ML1', 'AI/ML2'];
        const sheetNames = allSheetNames.filter(name => validSheetNames.includes(name));
        
        console.log('SRM Offline - All sheet names in workbook:', allSheetNames);
        console.log('SRM Offline - Valid sheet names found:', sheetNames);
        
        setAvailableSheets(sheetNames);
        
        // Process each sheet
        const allData = new Map<string, SRMOfflineStudent[]>();
        const dateSet = new Set<{ date: string; fullText: string }>();
        
        sheetNames.forEach(sheetName => {
          console.log(`SRM Offline - Processing sheet: ${sheetName}`);
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = processSRMOfflineSheet(worksheet, sheetName);
          console.log(`SRM Offline - Sheet ${sheetName} processed students:`, sheetData.length);
          console.log(`SRM Offline - First few students:`, sheetData.slice(0, 3));
          
          allData.set(sheetName, sheetData);
          
          // Extract dates from ANY student's attendance data
          if (sheetData.length > 0) {
            // Try to get dates from multiple students in case first one has no attendance data
            for (let i = 0; i < Math.min(10, sheetData.length); i++) {
              Object.keys(sheetData[i].attendance).forEach(date => {
                if (date && date.trim() !== '') {
                  dateSet.add({ date, fullText: `${date} - ${sheetName}` });
                }
              });
            }
            
            // Log the first few students for debugging
            console.log(`SRM Offline - First 3 students from ${sheetName}:`, sheetData.slice(0, 3).map(s => ({
              name: s.name,
              email: s.email,
              attendanceDates: Object.keys(s.attendance)
            })));
          }
        });
        
        console.log('SRM Offline - All processed data:', allData);
        console.log('SRM Offline - Extracted dates:', Array.from(dateSet));
        
        setAllSheetsData(allData);
        const sortedDates = Array.from(dateSet).sort((a, b) => {
          // Sort dates in descending order (newest first)
          const dateA = new Date(a.date.split('-').reverse().join('-'));
          const dateB = new Date(b.date.split('-').reverse().join('-'));
          return dateB.getTime() - dateA.getTime();
        });
        setAttendanceDates(sortedDates);
        
        if (sortedDates.length > 0) {
          setSelectedDate(sortedDates[0].date);
        }
        
        // Auto-select all sheets
        setSelectedSheetsForProcessing(new Set(sheetNames));
        setIsUploadComplete(true);
        
        console.log('SRM Offline - Processing complete:', {
          sheetsCount: sheetNames.length,
          datesCount: sortedDates.length,
          selectedDate: sortedDates[0]?.date,
          uploadComplete: true
        });
        
      } catch (error) {
        console.error('SRM Offline - Error processing attendance file:', error);
        console.error('SRM Offline - Error details:', error);
      } finally {
        setIsProcessing(false);
      }
    };
    
    reader.readAsArrayBuffer(file);
  };

  const processSRMOfflineSheet = (worksheet: XLSX.WorkSheet, sheetName: string): SRMOfflineStudent[] => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const students: SRMOfflineStudent[] = [];
    
    console.log(`SRM Offline - Processing sheet ${sheetName}:`, {
      totalRows: jsonData.length,
      headerRow: jsonData[0]
    });
    
    if (jsonData.length < 2) {
      console.log(`SRM Offline - Sheet ${sheetName} has insufficient data`);
      return students;
    }

    // Define row ranges for each batch (same as SRM Online for now)
    // Handle both naming conventions: AI/ML-1, AI/ML-2 and AI/ML1, AI/ML2
    const batchRowRanges: { [key: string]: { start: number; end: number } } = {
      'MS1': { start: 4, end: 135 },      // Row 5 to 135 (0-based: 4 to 135)
      'MS2': { start: 4, end: 114 },      // Row 5 to 114 (0-based: 4 to 114)
      'AI/ML-1': { start: 4, end: 104 },  // Row 5 to 104 (0-based: 4 to 104)
      'AI/ML-2': { start: 4, end: 111 },  // Row 5 to 111 (0-based: 4 to 111)
      'AI/ML1': { start: 4, end: 104 },   // Alternative naming: Row 5 to 104 (0-based: 4 to 104)
      'AI/ML2': { start: 4, end: 111 }    // Alternative naming: Row 5 to 111 (0-based: 4 to 111)
    };

    const rowRange = batchRowRanges[sheetName];
    if (!rowRange) {
      console.warn(`SRM Offline - Unknown sheet name: ${sheetName}, processing all available rows`);
    }
    
    // Get header row to find date columns
    const headerRow = jsonData[0] as unknown[];
    const dateColumns: { [key: number]: string } = {};
    
    console.log(`SRM Offline - Header row for ${sheetName}:`, headerRow);
    
    // Find date columns starting from column F (index 5)
    for (let i = 5; i < headerRow.length; i++) {
      const cellValue = headerRow[i];
      if (cellValue) {
        const cellStr = String(cellValue).trim();
        console.log(`SRM Offline - Header cell ${i}:`, cellStr);
        
        // Check if it's a date in various formats
        const dateMatch = cellStr.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/);
        
        // Also check for Excel serial date numbers (common in .xlsx files)
        if (!dateMatch && !isNaN(Number(cellStr))) {
          const excelDate = Number(cellStr);
          if (excelDate > 40000 && excelDate < 50000) { // Rough range for 2009-2037
            // Convert Excel serial date to JS date
            const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
            const day = String(jsDate.getDate()).padStart(2, '0');
            const month = String(jsDate.getMonth() + 1).padStart(2, '0');
            const year = jsDate.getFullYear();
            dateColumns[i] = `${day}-${month}-${year}`;
            console.log(`SRM Offline - Converted Excel date ${excelDate} to ${day}-${month}-${year}`);
          }
        } else if (dateMatch) {
          // Normalize date format to DD-MM-YYYY
          const normalizedDate = dateMatch[1].replace(/\//g, '-');
          dateColumns[i] = normalizedDate;
          console.log(`SRM Offline - Found date in header: ${normalizedDate}`);
        }
      }
    }
    
    console.log(`SRM Offline - Found date columns in ${sheetName}:`, dateColumns);
    
    // Process student data using batch-specific row ranges
    const startRow = rowRange ? rowRange.start : 1; // Default to row 2 (index 1) if no range specified
    const endRow = rowRange ? Math.min(rowRange.end + 1, jsonData.length) : jsonData.length;
    
    console.log(`SRM Offline - Processing rows ${startRow} to ${endRow} for ${sheetName}`);
    
    let processedCount = 0;
    for (let i = startRow; i < endRow; i++) {
      const row = jsonData[i] as unknown[];
      
      if (!row || row.length < 5) {
        console.log(`SRM Offline - Skipping row ${i}: insufficient columns`);
        continue;
      }
      
      const serialNo = row[0] ? String(row[0]).trim() : '';
      const name = row[1] ? String(row[1]).trim() : '';
      const email = row[2] ? String(row[2]).trim() : '';
      const regnNumber = row[3] ? String(row[3]).trim() : '';
      const contactNumber = row[4] ? String(row[4]).trim() : '';
      
      // Skip header rows and empty names, but be more lenient
      if (!name || name === '' || name.trim() === '' || 
          name.toLowerCase().includes('name') || 
          name.toLowerCase().includes('s.no') ||
          name.toLowerCase().includes('serial')) {
        console.log(`SRM Offline - Skipping row ${i}: invalid name "${name}"`);
        continue;
      }
      
      // Process attendance data
      const attendance: { [key: string]: number } = {};
      Object.keys(dateColumns).forEach(colIndex => {
        const date = dateColumns[parseInt(colIndex)];
        const attendanceValue = row[parseInt(colIndex)];
        
        if (attendanceValue !== undefined && attendanceValue !== null && attendanceValue !== '') {
          const numValue = parseInt(String(attendanceValue));
          attendance[date] = numValue === 1 ? 1 : 0;
        }
      });
      
      students.push({
        serialNo,
        name,
        email,
        regnNumber,
        contactNumber,
        attendance
      });
      
      processedCount++;
      if (processedCount <= 3) {
        console.log(`SRM Offline - Sample student ${processedCount}:`, {
          name,
          email,
          attendanceDates: Object.keys(attendance).length
        });
      }
    }
    
    console.log(`SRM Offline - Sheet ${sheetName} processed ${students.length} students`);
    return students;
  };

  const handleDateChange = (date: string) => {
    setSelectedDate(date);
  };

  // Email generation functions (similar to SRM Online but with SRM Offline branding)
  const generateAbsentStudentEmail = () => {
    if (!attendanceStats) return;
    
    const formatDateForEmail = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('-');
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const monthName = monthNames[parseInt(month) - 1];
      return `${day} ${monthName} ${year}`;
    };

    const formattedDate = formatDateForEmail(attendanceStats.date);
    const subjectLine = `Absent Students - MyAnatomy Training Session ${formattedDate} - SRM Offline`;
    
    const htmlContent = `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
<p><strong>Subject:</strong> Absent Students - MyAnatomy Training Session ${formattedDate}</p>

<p>Dear SRM Offline Faculty,</p>

<p>Greetings!</p>

<p>We would like to inform you about the students who were <strong>absent</strong> during today's MyAnatomy offline training session held on <strong>${formattedDate}</strong>.</p>

<p><strong>Session Details:</strong></p>
<ul>
<li><strong>Date:</strong> ${formattedDate}</li>
<li><strong>Platform:</strong> MyAnatomy Offline Campus</li>
<li><strong>Session Type:</strong> Interactive Offline Training Session</li>
<li><strong>Batches Covered:</strong> ${Array.from(selectedSheetsForProcessing).join(', ')}</li>
</ul>

<p><strong>Attendance Summary:</strong></p>
<ul>
<li><strong>Total Students:</strong> ${attendanceStats.totalStudents}</li>
<li><strong>Present:</strong> ${attendanceStats.present} (${attendanceStats.presentPercentage}%)</li>
<li><strong>Absent:</strong> ${attendanceStats.absent} (${attendanceStats.absentPercentage}%)</li>
</ul>

<p>The absent students have been BCC'd in this email for their awareness and necessary action.</p>

<p>We request your support in ensuring better attendance for upcoming offline sessions. Please encourage students to maintain regular attendance for maximum benefit from the MyAnatomy training program.</p>

<p>For any queries or support, please feel free to contact us at support@myanatomy.in</p>

<p><strong>Best regards,</strong><br>
MyAnatomy Team<br>
Digital Learning Solutions</p>
</div>`;

    const srmStaffEmails = 'offline.coordinator@srm.ac.in, dean@srm.ac.in, hod@srm.ac.in, admin@srm.ac.in';
    const myAnatomyStaffEmails = 'nishi.s@myanatomy.in, sucharita@myanatomy.in';
    
    let absentStudentEmails = '';
    if (attendanceStats && attendanceStats.absentStudents.length > 0) {
      absentStudentEmails = attendanceStats.absentStudents.map(student => student.email).filter(email => email).join(', ');
    }

    setAbsentStudentEmailContent(htmlContent);
    setAbsentStudentEmailSubject(subjectLine);
    setAbsentStudentEmailTo(srmStaffEmails);
    setAbsentStudentEmailCC(myAnatomyStaffEmails);
    setAbsentStudentEmailBCC(absentStudentEmails);
  };

  const generatePresentStudentEmail = () => {
    if (!attendanceStats) return;
    
    const formatDateForEmail = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('-');
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const monthName = monthNames[parseInt(month) - 1];
      return `${day} ${monthName} ${year}`;
    };

    const formattedDate = formatDateForEmail(attendanceStats.date);
    const subjectLine = `Present Students - MyAnatomy Training Session ${formattedDate} - SRM Offline`;
    
    const htmlContent = `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
<p><strong>Subject:</strong> Present Students - MyAnatomy Training Session ${formattedDate}</p>

<p>Dear SRM Offline Faculty,</p>

<p>Greetings!</p>

<p>We are pleased to inform you about the students who were <strong>present</strong> during today's MyAnatomy offline training session held on <strong>${formattedDate}</strong>.</p>

<p><strong>Session Details:</strong></p>
<ul>
<li><strong>Date:</strong> ${formattedDate}</li>
<li><strong>Platform:</strong> MyAnatomy Offline Campus</li>
<li><strong>Session Type:</strong> Interactive Offline Training Session</li>
<li><strong>Batches Covered:</strong> ${Array.from(selectedSheetsForProcessing).join(', ')}</li>
</ul>

<p><strong>Attendance Summary:</strong></p>
<ul>
<li><strong>Total Students:</strong> ${attendanceStats.totalStudents}</li>
<li><strong>Present:</strong> ${attendanceStats.present} (${attendanceStats.presentPercentage}%)</li>
<li><strong>Absent:</strong> ${attendanceStats.absent} (${attendanceStats.absentPercentage}%)</li>
</ul>

<p>The present students have been BCC'd in this email for their recognition and encouragement.</p>

<p>We appreciate the active participation of these students and look forward to their continued engagement in upcoming offline sessions.</p>

<p>For any queries or support, please feel free to contact us at support@myanatomy.in</p>

<p><strong>Best regards,</strong><br>
MyAnatomy Team<br>
Digital Learning Solutions</p>
</div>`;

    const srmStaffEmails = 'offline.coordinator@srm.ac.in, dean@srm.ac.in, hod@srm.ac.in, admin@srm.ac.in';
    const myAnatomyStaffEmails = 'nishi.s@myanatomy.in, sucharita@myanatomy.in';
    
    let presentStudentEmails = '';
    if (attendanceStats && attendanceStats.presentStudents.length > 0) {
      presentStudentEmails = attendanceStats.presentStudents.map(student => student.email).filter(email => email).join(', ');
    }

    setPresentStudentEmailContent(htmlContent);
    setPresentStudentEmailSubject(subjectLine);
    setPresentStudentEmailTo(srmStaffEmails);
    setPresentStudentEmailCC(myAnatomyStaffEmails);
    setPresentStudentEmailBCC(presentStudentEmails);
  };

  // Copy functions
  const copyAbsentStudentEmail = async () => {
    if (!absentStudentEmailContent) {
      generateAbsentStudentEmail();
      return;
    }

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([absentStudentEmailContent], { type: 'text/html' }),
          'text/plain': new Blob([absentStudentEmailContent.replace(/<[^>]*>/g, '')], { type: 'text/plain' })
        })
      ];
      await navigator.clipboard.write(clipboardData);
      setCopiedAbsentStudentEmail(true);
      setTimeout(() => setCopiedAbsentStudentEmail(false), 2000);
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
          'text/plain': new Blob([presentStudentEmailContent.replace(/<[^>]*>/g, '')], { type: 'text/plain' })
        })
      ];
      await navigator.clipboard.write(clipboardData);
      setCopiedPresentStudentEmail(true);
      setTimeout(() => setCopiedPresentStudentEmail(false), 2000);
    } catch (error) {
      console.error('Failed to copy present student email:', error);
    }
  };

  const copyAbsentEmails = async () => {
    if (!attendanceStats) return;
    const emails = attendanceStats.absentStudents.map(s => s.email).filter(email => email).join(', ');
    try {
      await navigator.clipboard.writeText(emails);
      setCopiedAbsentEmails(true);
      setTimeout(() => setCopiedAbsentEmails(false), 2000);
    } catch (error) {
      console.error('Failed to copy absent emails:', error);
    }
  };

  const copyPresentEmails = async () => {
    if (!attendanceStats) return;
    const emails = attendanceStats.presentStudents.map(s => s.email).filter(email => email).join(', ');
    try {
      await navigator.clipboard.writeText(emails);
      setCopiedPresentEmails(true);
      setTimeout(() => setCopiedPresentEmails(false), 2000);
    } catch (error) {
      console.error('Failed to copy present emails:', error);
    }
  };

  const generateEmailTemplate = () => {
    if (!attendanceStats || !selectedDate) return;
    
    const formatDate = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('-');
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const monthName = monthNames[parseInt(month) - 1];
      return `${day} ${monthName} ${year}`;
    };

    const formattedDate = formatDate(selectedDate);
    const selectedBatches = Array.from(selectedSheetsForProcessing);

    const content = `Subject: SRM Offline - MyAnatomy Training Session Report - ${formattedDate}

Dear SRM Offline Faculty,

Greetings!

We are pleased to share the attendance report for the MyAnatomy offline training session conducted on ${formattedDate}.

Session Details:
- Date: ${formattedDate}
- Platform: MyAnatomy Offline Campus
- Mode: Offline Training Session
- Batches Covered: ${selectedBatches.join(', ')}

Attendance Summary:
- Total Students: ${attendanceStats.totalStudents}
- Present: ${attendanceStats.present} (${attendanceStats.presentPercentage}%)
- Absent: ${attendanceStats.absent} (${attendanceStats.absentPercentage}%)

Training Materials and Session Records:
${emailTemplate.sheetsLink}

The session covered comprehensive anatomy modules with hands-on practical activities. Students who attended the session have access to recorded materials and additional resources.

For any queries regarding the session or student attendance, please feel free to contact us.

Best regards,
MyAnatomy Team
Digital Learning Solutions

---
📧 Generated with MyAnatomy Attendance Compiler
🔗 Session Materials: ${emailTemplate.sheetsLink}`;

    setEmailTemplate(prev => ({ 
      ...prev, 
      trainingDate: formattedDate,
      batches: selectedBatches,
      generatedContent: content 
    }));
  };

  const copyEmailTemplate = async () => {
    try {
      await navigator.clipboard.writeText(emailTemplate.generatedContent);
      setCopiedEmailTemplate(true);
      setTimeout(() => setCopiedEmailTemplate(false), 2000);
    } catch (error) {
      console.error('Failed to copy email template:', error);
    }
  };

  const copyEmailTo = async () => {
    try {
      await navigator.clipboard.writeText(emailTemplate.to);
      setCopiedEmailTo(true);
      setTimeout(() => setCopiedEmailTo(false), 2000);
    } catch (error) {
      console.error('Failed to copy email to:', error);
    }
  };

  return (
    <div className="space-y-8">
      {/* Row 1: Upload and Student Selection - Same structure as NIET/SRM Online */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Box 1: Upload Document Section */}
        <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
          <div className="flex items-center gap-3 mb-4">
            <Upload className="w-5 h-5 text-gray-400" />
            <h2 className="text-lg font-semibold text-white">Upload Files</h2>
            <span className="bg-purple-600 text-white text-xs px-2 py-1 rounded-full">
              SRM Offline
            </span>
          </div>
          
          <label htmlFor="attendance-sheet" className="text-sm text-gray-400 mb-2 block">Attendance Sheet</label>
          <div className="space-y-4">
            <div className="relative border-2 border-dashed border-purple-600 rounded-lg p-8 text-center bg-purple-600/10">
              <div className="flex flex-col items-center justify-center">
                <FileText className="w-10 h-10 text-purple-500 mb-3" />
                <p className="text-white font-medium mb-2">SRM Offline Attendance Sheet</p>
                <p className="text-xs text-gray-400 mb-4">
                  {attendanceFile ? `Loaded: ${attendanceFile.name}` : 'Use the pre-configured attendance sheet or upload a custom one'}
                </p>
                <button 
                  onClick={loadSRMOfflineAttendanceSheet}
                  className="px-4 py-2 bg-purple-600 text-white rounded-md hover:bg-purple-700 transition-colors font-medium"
                >
                  Load SRM Offline Attendance Sheet
                </button>
              </div>
            </div>

            <div className="text-center">
              <p className="text-xs text-gray-400 mb-2">or upload your own</p>
              <label className="cursor-pointer px-4 py-2 bg-gray-600 text-white rounded-md hover:bg-gray-700 transition-colors font-medium inline-block">
                <Upload className="w-4 h-4 inline mr-2" />
                Upload Excel File
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                />
              </label>
            </div>

            {isProcessing && (
              <div className="text-center text-purple-400">
                <div className="inline-block animate-spin rounded-full h-6 w-6 border-b-2 border-purple-400 mr-2"></div>
                Processing attendance data...
              </div>
            )}
          </div>

          {(allSheetsData.size > 0 || attendanceDates.length > 0) && (
            <div className="mt-4 p-4 bg-gray-700/50 rounded-lg">
              <h3 className="text-sm font-medium text-white mb-2">
                Attendance Data Loaded ({attendanceDates.length} dates found)
                {availableSheets.length > 0 && (
                  <span className="ml-2 text-xs bg-gray-600 text-gray-300 px-2 py-0.5 rounded">
                    Batches: {availableSheets.join(', ')}
                  </span>
                )}
              </h3>
              <div className="max-h-32 overflow-y-auto">
                {attendanceDates.slice(0, 5).map((dateObj, index) => (
                  <div key={index} className="text-xs text-gray-300 py-1">
                    {dateObj.date} - {dateObj.fullText}
                  </div>
                ))}
                {attendanceDates.length > 5 && (
                  <div className="text-xs text-gray-400">...and {attendanceDates.length - 5} more dates</div>
                )}
              </div>
            </div>
          )}
        </section>

        {/* Box 2: Attendance Analysis - Same structure as SRM Online */}
        <section className={`bg-gray-800/50 border border-gray-700/50 rounded-xl p-6 transition-opacity ${!isUploadComplete ? 'opacity-50' : 'opacity-100'}`}>
          <div className="flex items-center gap-3 mb-4">
            <Book className="w-5 h-5 text-gray-400" />
            <h2 className="text-lg font-semibold text-white">Attendance Analysis</h2>
          </div>
          <p className="text-sm text-gray-400 mb-4">Select a date to view attendance statistics.</p>
          
          {isUploadComplete && attendanceDates.length > 0 ? (
            <div className="space-y-4">
              <div className="flex justify-between items-center">
                <span className="text-sm text-gray-300">
                  {attendanceDates.length} attendance dates available
                </span>
                <div className="flex gap-2">
                  <span className="text-xs px-2 py-1 bg-purple-600 text-white rounded">
                    SRM Offline
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
                  className="w-full bg-gray-700 border border-gray-600 rounded-md px-3 py-2 text-white focus:ring-2 focus:ring-purple-500 focus:border-purple-500 select-auto"
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
                  
                  <div className="bg-purple-600/20 border border-purple-500/30 rounded-lg p-3">
                    <div className="text-purple-300 text-sm font-medium">Total Students</div>
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
                  
                  {/* All Batches Summary */}
                  {allSheetsAttendanceData.size > 0 && (
                    <div className="mt-6 bg-gray-600/30 rounded-lg p-4">
                      <h4 className="text-sm font-semibold text-white mb-3">
                        All Batches Summary for {attendanceStats.date}
                      </h4>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                        {Array.from(allSheetsAttendanceData.entries()).map(([batchName, stats]) => (
                          <div key={batchName} className="bg-gray-700/50 rounded p-3">
                            <div className="text-xs font-medium text-purple-300 mb-2">{batchName}</div>
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

                  {/* Present/Absent Student Lists with Copy */}
                  <div className="mt-6 grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Present Students */}
                    <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
                      <div className="flex items-center justify-between mb-3">
                        <div>
                          <h4 className="text-green-300 text-sm font-semibold">
                            Present Students ({attendanceStats.presentStudents.length})
                          </h4>
                          <p className="text-xs text-green-400/70 mt-1">
                            📧 Ready for Gmail &quot;BCC&quot; field
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
                            📧 Ready for Gmail &quot;BCC&quot; field
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
            <div className="text-center py-10">
              <Book className="w-12 h-12 text-gray-600 mx-auto mb-3" />
              <p className="text-gray-500">Upload a file to view attendance statistics...</p>
            </div>
          )}
        </section>
      </div>

      {/* Row 2: Email Template Generator - Same structure as SRM Online */}
      {attendanceStats && (
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
            <div className="flex items-center gap-3 mb-4">
              <Mail className="w-5 h-5 text-gray-400" />
              <h2 className="text-lg font-semibold text-white">Email Template</h2>
            </div>
            <div className="bg-gray-700/50 rounded-lg p-4">
              <h3 className="text-sm font-medium text-white mb-3">Generate Training Report Email</h3>
              
              <div className="space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-xs text-gray-300 font-medium">Training Date:</label>
                    <input
                      type="text"
                      value={emailTemplate.trainingDate}
                      onChange={(e) => setEmailTemplate(prev => ({ ...prev, trainingDate: e.target.value }))}
                      className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
                      placeholder="e.g., 28 August 2025"
                    />
                  </div>
                  <div>
                    <label className="text-xs text-gray-300 font-medium">Sheets Link:</label>
                    <input
                      type="text"
                      value={emailTemplate.sheetsLink}
                      onChange={(e) => setEmailTemplate(prev => ({ ...prev, sheetsLink: e.target.value }))}
                      className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
                    />
                  </div>
                </div>
                
                <div>
                  <label className="text-xs text-gray-300 font-medium">To (Email Recipients):</label>
                  <div className="flex gap-2">
                    <textarea
                      value={emailTemplate.to}
                      onChange={(e) => setEmailTemplate(prev => ({ ...prev, to: e.target.value }))}
                      placeholder="offline.coordinator@srm.ac.in, dean@srm.ac.in, ..."
                      className="flex-1 bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-purple-500 focus:border-purple-500 resize-none h-16"
                    />
                    <button
                      onClick={copyEmailTo}
                      className="flex items-center justify-center w-10 h-10 bg-gray-600 hover:bg-gray-700 text-white rounded transition-colors"
                      title="Copy recipients"
                    >
                      {copiedEmailTo ? (
                        <Check className="w-4 h-4" />
                      ) : (
                        <Copy className="w-4 h-4" />
                      )}
                    </button>
                  </div>
                </div>
                
                <div className="flex gap-2">
                  <button
                    onClick={generateEmailTemplate}
                    className="px-4 py-2 bg-purple-600 text-white font-medium rounded-md hover:bg-purple-700 transition-colors"
                  >
                    Generate Template
                  </button>
                  
                  {emailTemplate.generatedContent && (
                    <button
                      onClick={copyEmailTemplate}
                      className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white font-medium rounded-md hover:bg-gray-700 transition-colors"
                    >
                      {copiedEmailTemplate ? (
                        <>
                          <Check className="w-4 h-4" />
                          Copied!
                        </>
                      ) : (
                        <>
                          <Copy className="w-4 h-4" />
                          Copy Template
                        </>
                      )}
                    </button>
                  )}
                </div>
                
                {emailTemplate.generatedContent && (
                  <div className="bg-gray-800 rounded p-3 max-h-64 overflow-y-auto">
                    <pre className="text-xs text-gray-300 whitespace-pre-wrap font-mono">
                      {emailTemplate.generatedContent}
                    </pre>
                  </div>
                )}
              </div>
            </div>
          </section>

          {/* Intern Report Section - Same structure as SRM Online */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
            <div className="flex items-center gap-3 mb-4">
              <ClipboardList className="w-5 h-5 text-gray-400" />
              <h2 className="text-lg font-semibold text-white">Intern Report</h2>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="text-xs text-gray-300 font-medium mb-2 block">
                  Session Topics (Enter topics covered during the session):
                </label>
                <textarea
                  value={internReport}
                  onChange={(e) => setInternReport(e.target.value)}
                  placeholder="Enter the topics and activities covered during the SRM Offline training session..."
                  className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-purple-500 focus:border-purple-500 resize-none"
                  rows={internReportExpanded ? 8 : 4}
                />
                <button
                  onClick={() => setInternReportExpanded(!internReportExpanded)}
                  className="text-xs text-purple-400 hover:text-purple-300 mt-1"
                >
                  {internReportExpanded ? 'Collapse' : 'Expand'} text area
                </button>
              </div>
              
              {internReport && (
                <div className="bg-gray-700/50 rounded-lg p-4">
                  <h4 className="text-sm font-medium text-white mb-3">Topics Preview</h4>
                  <div className="text-xs text-gray-300 whitespace-pre-wrap max-h-32 overflow-y-auto">
                    {internReport.split('\n').map((line, index) => (
                      <div key={index} className="mb-1">
                        <strong>{index + 1}.</strong> {line}
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </section>
        </div>
      )}

      {/* Row 3: Generated Emails - Same structure as SRM Online */}
      {attendanceStats && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Absent Students Email */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
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
                  className="flex items-center gap-1 px-3 py-1.5 bg-gray-600 hover:bg-gray-700 text-white text-xs font-medium rounded-md transition-colors disabled:opacity-50"
                >
                  {copiedAbsentStudentEmail ? (
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
            </div>
            
            {absentStudentEmailContent && (
              <div className="bg-gray-700/50 rounded p-3 max-h-64 overflow-y-auto">
                <div 
                  className="text-xs text-gray-300 leading-relaxed"
                  dangerouslySetInnerHTML={{ __html: absentStudentEmailContent }}
                />
              </div>
            )}
          </section>

          {/* Present Students Email */}
          <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
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
                  className="flex items-center gap-1 px-3 py-1.5 bg-gray-600 hover:bg-gray-700 text-white text-xs font-medium rounded-md transition-colors disabled:opacity-50"
                >
                  {copiedPresentStudentEmail ? (
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
            </div>
            
            {presentStudentEmailContent && (
              <div className="bg-gray-700/50 rounded p-3 max-h-64 overflow-y-auto">
                <div 
                  className="text-xs text-gray-300 leading-relaxed"
                  dangerouslySetInnerHTML={{ __html: presentStudentEmailContent }}
                />
              </div>
            )}
          </section>
        </div>
      )}
    </div>
  );
}