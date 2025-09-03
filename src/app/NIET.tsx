'use client';

import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { FileText, Upload, Book, Copy, Check, Mail, ChevronDown, Users } from 'lucide-react';

interface NIETStudent {
  serialNo: string;
  rollNumber: string;
  name: string;
  email: string;
  attendance: { [key: string]: number }; // date -> 0/1
}

interface NIETAttendanceStats {
  date: string;
  totalStudents: number;
  present: number;
  absent: number;
  presentPercentage: number;
  absentPercentage: number;
  presentStudents: Array<{ name: string; email: string; rollNumber: string }>;
  absentStudents: Array<{ name: string; email: string; rollNumber: string }>;
}

interface EmailTemplate {
  trainingDate: string;
  batches: string[];
  sheetsLink: string;
  to: string;
  cc: string;
  generatedContent: string;
}

interface NIETProps {
  isVisible: boolean;
}

export default function NIET({ isVisible }: NIETProps) {
  // File and data states
  const [attendanceFile, setAttendanceFile] = useState<File | null>(null);
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>(''); // Primary sheet for date selection
  const [selectedSheetsForProcessing, setSelectedSheetsForProcessing] = useState<Set<string>>(new Set());
  const [selectedEmailBatchSheet, setSelectedEmailBatchSheet] = useState<string>(''); // For student email templates
  const [allSheetsData, setAllSheetsData] = useState<Map<string, NIETStudent[]>>(new Map());
  const [allSheetsAttendanceData, setAllSheetsAttendanceData] = useState<Map<string, NIETAttendanceStats>>(new Map());
  const [attendanceDates, setAttendanceDates] = useState<Array<{ date: string; fullText: string }>>([]);
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [attendanceStats, setAttendanceStats] = useState<NIETAttendanceStats | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploadComplete, setIsUploadComplete] = useState(false);

  // Email states
  const [absentStudentEmailContent, setAbsentStudentEmailContent] = useState<string>('');
  const [absentStudentEmailContentForCopy, setAbsentStudentEmailContentForCopy] = useState<string>('');
  const [presentStudentEmailContent, setPresentStudentEmailContent] = useState<string>('');
  const [presentStudentEmailContentForCopy, setPresentStudentEmailContentForCopy] = useState<string>('');
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
    sheetsLink: 'https://docs.google.com/spreadsheets/d/1X3p75fw2Uz34A-LveIhex-zAxm-VVlImNOmoZOihvMg/edit?gid=1842156615#gid=1842156615',
    to: 'premsagar.sharma@niet.co.in, amd@niet.co.in, director@niet.co.in, arvind.sharma@niet.co.in, ajeet.singh@niet.co.in',
    cc: 'nishi.s@myanatomy.in, sucharita@myanatomy.in',
    generatedContent: ''
  });
  const [copiedEmailTemplate, setCopiedEmailTemplate] = useState<boolean>(false);
  const [copiedEmailTo, setCopiedEmailTo] = useState<boolean>(false);
  const [copiedEmailCc, setCopiedEmailCc] = useState<boolean>(false);
  const [emailTemplateSubject, setEmailTemplateSubject] = useState<string>('');
  const [copiedEmailTemplateSubject, setCopiedEmailTemplateSubject] = useState<boolean>(false);
  const [copiedAbsentStudentSubject, setCopiedAbsentStudentSubject] = useState<boolean>(false);
  const [copiedPresentStudentSubject, setCopiedPresentStudentSubject] = useState<boolean>(false);
  const [copiedAbsentStudentTo, setCopiedAbsentStudentTo] = useState<boolean>(false);
  const [copiedAbsentStudentCc, setCopiedAbsentStudentCc] = useState<boolean>(false);
  const [copiedAbsentStudentBcc, setCopiedAbsentStudentBcc] = useState<boolean>(false);
  const [copiedPresentStudentTo, setCopiedPresentStudentTo] = useState<boolean>(false);
  const [copiedPresentStudentCc, setCopiedPresentStudentCc] = useState<boolean>(false);
  const [copiedPresentStudentBcc, setCopiedPresentStudentBcc] = useState<boolean>(false);

  // Intern report states
  const [internReport, setInternReport] = useState<string>('');
  const [internReportExpanded, setInternReportExpanded] = useState<boolean>(false);

  const calculateAttendanceStats = useCallback((date: string): NIETAttendanceStats | null => {
    console.log('CALC STATS DEBUG - Starting calculation for PRIMARY BATCH ONLY:', {
      date,
      primarySheet: selectedSheet,
      allSheetsDataCount: allSheetsData.size
    });
    
    if (!selectedSheet || !allSheetsData.size) {
      console.log('CALC STATS DEBUG - Early return: no primary sheet or no data');
      return null;
    }
    
    // Get students ONLY from the primary batch (selectedSheet)
    const primarySheetData = allSheetsData.get(selectedSheet);
    if (!primarySheetData) {
      console.log('CALC STATS DEBUG - Early return: no data for primary sheet');
      return null;
    }
    
    console.log(`CALC STATS DEBUG - Using ONLY students from primary batch: ${selectedSheet} (${primarySheetData.length} students)`);
    
    if (primarySheetData.length === 0) {
      console.log('CALC STATS DEBUG - Early return: no students in primary sheet');
      return null;
    }
    
    const presentStudents: Array<{ name: string; email: string; rollNumber: string }> = [];
    const absentStudents: Array<{ name: string; email: string; rollNumber: string }> = [];
    
    // Process students from PRIMARY BATCH ONLY
    primarySheetData.forEach(student => {
      const attendanceValue = student.attendance[date];
      if (attendanceValue === 1) {
        presentStudents.push({
          name: student.name,
          email: student.email,
          rollNumber: student.rollNumber
        });
      } else if (attendanceValue === 0) {
        absentStudents.push({
          name: student.name,
          email: student.email,
          rollNumber: student.rollNumber
        });
      }
    });
    
    console.log('CALC STATS DEBUG - Final counts from PRIMARY BATCH ONLY:', {
      primaryBatch: selectedSheet,
      presentStudents: presentStudents.length,
      absentStudents: absentStudents.length,
      totalStudents: presentStudents.length + absentStudents.length
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
  }, [selectedSheet, allSheetsData]);

  useEffect(() => {
    if (selectedDate && selectedSheet) {
      const stats = calculateAttendanceStats(selectedDate);
      setAttendanceStats(stats);
      
      // Calculate stats for all sheets
      const allSheetStats = new Map<string, NIETAttendanceStats>();
      selectedSheetsForProcessing.forEach(sheetName => {
        const sheetData = allSheetsData.get(sheetName);
        if (sheetData) {
          const presentStudents = sheetData.filter(student => student.attendance[selectedDate] === 1)
            .map(student => ({ name: student.name, email: student.email, rollNumber: student.rollNumber }));
          const absentStudents = sheetData.filter(student => student.attendance[selectedDate] === 0)
            .map(student => ({ name: student.name, email: student.email, rollNumber: student.rollNumber }));
          
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
  }, [selectedDate, selectedSheet, selectedSheetsForProcessing, allSheetsData, calculateAttendanceStats]);

  if (!isVisible) return null;

  // Functions
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
      processAttendanceFile(file);
    } catch (error) {
      console.error('Failed to load NIET attendance sheet:', error);
    }
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    console.log('FILE UPLOAD DEBUG - File selected:', {
      fileName: file?.name,
      fileSize: file?.size,
      fileType: file?.type,
      isXlsx: file?.name.endsWith('.xlsx')
    });
    
    if (file && file.name.endsWith('.xlsx')) {
      console.log('FILE UPLOAD DEBUG - Valid Excel file, setting state and processing');
      setAttendanceFile(file);
      processAttendanceFile(file);
    } else {
      console.log('FILE UPLOAD DEBUG - Invalid file or no file selected');
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
        
        console.log('Workbook sheet names:', workbook.SheetNames);
        
        // Get available sheets and filter to only include the 4 specified batch names
        const allSheetNames = workbook.SheetNames;
        const validSheetNames = ['MS1 Attendance', 'MS-1 Attendance', 'JAVA SDE-1 Attendance', 'JAVA SDE-2 Attendance', 'Data Scientist Attendance'];
        const sheetNames = allSheetNames.filter(name => validSheetNames.map(n => n.toLowerCase().trim()).includes(name.toLowerCase().trim()));
        
        console.log('All sheet names in workbook:', allSheetNames);
        console.log('Valid NIET sheet names found:', sheetNames);
        
        setAvailableSheets(sheetNames);
        
        // Set primary sheet (first available sheet for date selection)
        if (sheetNames.length > 0) {
          setSelectedSheet(sheetNames[0]);
          console.log('FINAL DEBUG - Set primary sheet:', sheetNames[0]);
          
          // Auto-select the primary sheet for processing
          setSelectedSheetsForProcessing(new Set([sheetNames[0]]));
          console.log('FINAL DEBUG - Auto-selected primary sheet for processing:', sheetNames[0]);
        }
        
        // Process each sheet
        const allData = new Map<string, NIETStudent[]>();
        const dateMap = new Map<string, { date: string; fullText: string }>(); // Use Map to prevent duplicates by date
        
        sheetNames.forEach(sheetName => {
          console.log(`Processing sheet: ${sheetName}`);
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = processNIETSheet(worksheet, sheetName);
          console.log(`Sheet ${sheetName} processed students:`, sheetData.length);
          console.log(`First few students:`, sheetData.slice(0, 3));
          
          allData.set(sheetName, sheetData);
          
          // Extract dates only from the primary sheet (first sheet)
          if (sheetData.length > 0 && sheetName === sheetNames[0]) {
            console.log(`Extracting dates from PRIMARY sheet: ${sheetName}`);
            // Try to get dates from multiple students in case first one has no attendance data
            for (let i = 0; i < Math.min(10, sheetData.length); i++) {
              Object.keys(sheetData[i].attendance).forEach(date => {
                if (date && date.trim() !== '') {
                  // Store dates from primary sheet only
                  if (!dateMap.has(date)) {
                    dateMap.set(date, { date, fullText: `${date}` });
                  }
                }
              });
            }
            
            // Log the first few students for debugging
            console.log(`First 3 students from ${sheetName}:`, sheetData.slice(0, 3).map(s => ({
              name: s.name,
              email: s.email,
              attendanceDates: Object.keys(s.attendance)
            })));
          }
        });
        
        console.log('All processed data:', allData);
        console.log('Extracted dates:', Array.from(dateMap.values()));
        
        console.log('FINAL DEBUG - Setting data states:', {
          allDataSize: allData.size,
          dateMapSize: dateMap.size,
          allDataKeys: Array.from(allData.keys()),
          allDatesFound: Array.from(dateMap.values()).map(d => d.date)
        });

        setAllSheetsData(allData);
        const sortedDates = Array.from(dateMap.values()).sort((a, b) => {
          // Sort dates in descending order (newest first)
          const dateA = new Date(a.date.split('-').reverse().join('-'));
          const dateB = new Date(b.date.split('-').reverse().join('-'));
          return dateB.getTime() - dateA.getTime();
        });
        
        console.log('FINAL DEBUG - Setting attendance dates:', sortedDates);
        setAttendanceDates(sortedDates);
        
        if (sortedDates.length > 0) {
          console.log('FINAL DEBUG - Setting selected date:', sortedDates[0].date);
          setSelectedDate(sortedDates[0].date);
        }
        
        // Don't auto-select sheets - let user choose
        console.log('FINAL DEBUG - Available sheets for selection:', sheetNames);
        setSelectedSheetsForProcessing(new Set()); // Start with no sheets selected
        
        console.log('FINAL DEBUG - Setting upload complete to true');
        setIsUploadComplete(true);
        
        console.log('Processing complete:', {
          sheetsCount: sheetNames.length,
          datesCount: sortedDates.length,
          selectedDate: sortedDates[0]?.date,
          uploadComplete: true
        });
        
      } catch (error) {
        console.error('Error processing attendance file:', error);
        console.error('Error details:', error);
      } finally {
        setIsProcessing(false);
      }
    };
    
    reader.readAsArrayBuffer(file);
  };

  const processNIETSheet = (worksheet: XLSX.WorkSheet, sheetName: string): NIETStudent[] => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const students: NIETStudent[] = [];
    
    console.log(`Processing sheet ${sheetName}:`, {
      totalRows: jsonData.length,
      headerRow: jsonData[0]
    });
    
    if (jsonData.length < 2) {
      console.log(`Sheet ${sheetName} has insufficient data`);
      return students;
    }

    // Get header row to find date columns (header is in row 4, index 3)
    const headerRow = jsonData[3] as unknown[];
    const dateColumns: { [key: number]: string } = {};
    
    console.log(`Header row for ${sheetName} (Row 4):`, headerRow);
    
    // Find date columns starting from column H (index 7)
    for (let i = 7; i < headerRow.length; i++) {
      const cellValue = headerRow[i];
      if (cellValue) {
        const cellStr = String(cellValue).trim();
        console.log(`Header cell ${i}:`, cellStr);
        
        // Extract DD/MM/YYYY from DD/MM/YYYY(Time) format
        const dateMatch = cellStr.match(/(\d{2}\/\d{2}\/\d{4})/);
        
        if (dateMatch) {
          const normalizedDate = dateMatch[1].replace(/\//g, '-'); // Convert to DD-MM-YYYY
          dateColumns[i] = normalizedDate;
          console.log(`Found date in header: ${normalizedDate} (from: ${cellStr})`);
        }
      }
    }
    
    console.log(`Found date columns in ${sheetName}:`, dateColumns);
    
    // Process student data starting from row 5 (index 4)
    const startRow = 4; 
    const endRow = jsonData.length;
    
    console.log(`Processing rows ${startRow} to ${endRow} for ${sheetName}`);
    
    let processedCount = 0;
    for (let i = startRow; i < endRow; i++) {
      const row = jsonData[i] as unknown[];
      
      if (!row || row.length < 8) { // Ensure enough columns for S.No, Roll, Name, Email, and at least one date
        console.log(`Skipping row ${i}: insufficient columns`);
        continue;
      }
      
      const serialNo = row[0] ? String(row[0]).trim() : '';
      const rollNumber = row[1] ? String(row[1]).trim() : '';
      const name = row[2] ? String(row[2]).trim() : '';
      const email = row[3] ? String(row[3]).trim() : '';
      
      // Skip header rows and empty names
      if (!name || name === '' || name.trim() === '' || 
          name.toLowerCase().includes('name') || 
          name.toLowerCase().includes('s.no') ||
          name.toLowerCase().includes('serial')) {
        console.log(`Skipping row ${i}: invalid name "${name}"`);
        continue;
      }
      
      // Process attendance data
      const attendance: { [key: string]: number } = {};
      Object.keys(dateColumns).forEach(colIndex => {
        const date = dateColumns[parseInt(colIndex)];
        const attendanceValue = row[parseInt(colIndex)];
        
        if (attendanceValue !== undefined && attendanceValue !== null && attendanceValue !== '') {
          const numValue = parseInt(String(attendanceValue));
          attendance[date] = numValue === 1 ? 1 : 0; // 0 for absent, 1 for present
        }
      });
      
      students.push({
        serialNo,
        rollNumber,
        name,
        email,
        attendance
      });
      
      processedCount++;
      if (processedCount <= 3) {
        console.log(`Sample student ${processedCount}:`, {
          name,
          email,
          attendanceDates: Object.keys(attendance).length
        });
      }
    }
    
    console.log(`Sheet ${sheetName} processed ${students.length} students`);
    return students;
  };

  const handleDateChange = (date: string) => {
    setSelectedDate(date);
  };

  // Batch selection functions (similar to NIET)
  const handleSheetSelectionToggle = (sheetName: string) => {
    const newSelection = new Set(selectedSheetsForProcessing);
    if (newSelection.has(sheetName)) {
      newSelection.delete(sheetName);
    } else {
      newSelection.add(sheetName);
    }
    setSelectedSheetsForProcessing(newSelection);
  };

  const selectAllSheets = () => {
    setSelectedSheetsForProcessing(new Set(availableSheets));
  };

  const clearAllSheets = () => {
    setSelectedSheetsForProcessing(new Set());
  };

  // Primary sheet change handler
  const handleSheetChange = (sheetName: string) => {
    setSelectedSheet(sheetName);
    
    // Auto-select the new primary sheet for processing
    const newSelection = new Set(selectedSheetsForProcessing);
    newSelection.add(sheetName);
    setSelectedSheetsForProcessing(newSelection);
    
    // Reprocess dates from the new primary sheet
    if (attendanceFile) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'array' });
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = processNIETSheet(worksheet, sheetName);
          
          // Extract dates from the new primary sheet
          const dateMap = new Map<string, { date: string; fullText: string }>();
          if (sheetData.length > 0) {
            for (let i = 0; i < Math.min(10, sheetData.length); i++) {
              Object.keys(sheetData[i].attendance).forEach(date => {
                if (date && date.trim() !== '') {
                  if (!dateMap.has(date)) {
                    dateMap.set(date, { date, fullText: `${date}` });
                  }
                }
              });
            }
          }
          
          const sortedDates = Array.from(dateMap.values()).sort((a, b) => {
            const dateA = new Date(a.date.split('-').reverse().join('-'));
            const dateB = new Date(b.date.split('-').reverse().join('-'));
            return dateB.getTime() - dateA.getTime();
          });
          
          setAttendanceDates(sortedDates);
          if (sortedDates.length > 0) {
            setSelectedDate(sortedDates[0].date);
          }
          
          console.log('PRIMARY SHEET CHANGED - New dates:', sortedDates.map(d => d.date));
        } catch (error) {
          console.error('Error reprocessing dates for new primary sheet:', error);
        }
      };
      reader.readAsArrayBuffer(attendanceFile);
    }
  };

  useEffect(() => {
    if (selectedEmailBatchSheet) {
      generateAbsentStudentEmail();
      generatePresentStudentEmail();
    }
  }, [selectedEmailBatchSheet, internReport, allSheetsAttendanceData, selectedDate]);

  // Email generation functions
  const generateAbsentStudentEmail = useCallback(() => {
    if (!selectedEmailBatchSheet || !allSheetsAttendanceData.has(selectedEmailBatchSheet)) return;

    const batchStats = allSheetsAttendanceData.get(selectedEmailBatchSheet);
    if (!batchStats) return; // Should not happen due to has() check, but for type safety
    
    const formatDateForEmail = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('-');
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const monthName = monthNames[parseInt(month) - 1];
      return `${day} ${monthName} ${year}`;
    };

    const formattedDate = formatDateForEmail(batchStats.date);
    
    // Map sheet names to proper topic names for subject line
    const getTopicName = (sheetName: string): string => {
      const name = sheetName.toLowerCase();
      if (name.includes('ms-1') || name.includes('mern')) {
        return 'MERN Stack';
      } else if (name.includes('java sde-1') || name.includes('java sde 1')) {
        return 'JAVA SDE-1';
      } else if (name.includes('java sde-2') || name.includes('java sde 2')) {
        return 'JAVA SDE-2';
      } else if (name.includes('data scientist') || name.includes('python')) {
        return 'Data Science Python';
      }
      return sheetName; // fallback to original name
    };
    
    const topicName = getTopicName(selectedEmailBatchSheet);
    const subjectLine = `${topicName} NCET + Online Training NIET College Attendance ${formattedDate}`;
    
    // Version for display (light text for dark UI)
    const htmlContentForDisplay = `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #E5E7EB;">
<p><strong>Dear Students</strong>,</p>

<p>This email is to address the issue of student attendance at our live training sessions. We have observed that some students have missed the live training session on <strong>${formattedDate}</strong>.</p>


<p>Here's what you missed during the session:</p>

${internReport ? formatInternReportToTable(internReport) : '<p><em>Session summary will be added here based on the intern report content.</em></p>'}

<p>We understand that unforeseen circumstances may arise, however, it is crucial to attend these sessions regularly. These live sessions are an integral part of your learning journey and provide valuable opportunities for interactive learning, Q&A, and engagement with instructors and fellow students.</p>

<p>Missing these free sessions is not only detrimental to your learning but also disrespectful to the instructors and other students who are diligently participating.</p>

<p>Students who continue to remain absent for sessions will be flagged, and appropriate escalations will be made with the Training and Placement Officers (TPOs) if this behaviour is continued.</p>

<p>We expect all students to attend all upcoming live training sessions promptly.</p>

<p>We urge you to prioritize your attendance and actively participate in these valuable sessions.</p>

<p>Regards,</p>
</div>`;

    // Version for copying (black text for email clients)
    const htmlContentForCopy = `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #000000;">
<p><strong>Dear Students</strong>,</p>

<p>This email is to address the issue of student attendance at our live training sessions. We have observed that some students have missed the live training session on <strong>${formattedDate}</strong>.</p>


<p>Here's what you missed during the session:</p>

${internReport ? formatInternReportToTable(internReport) : '<p><em>Session summary will be added here based on the intern report content.</em></p>'}

<p>We understand that unforeseen circumstances may arise, however, it is crucial to attend these sessions regularly. These live sessions are an integral part of your learning journey and provide valuable opportunities for interactive learning, Q&A, and engagement with instructors and fellow students.</p>

<p>Missing these free sessions is not only detrimental to your learning but also disrespectful to the instructors and other students who are diligently participating.</p>

<p>Students who continue to remain absent for sessions will be flagged, and appropriate escalations will be made with the Training and Placement Officers (TPOs) if this behaviour is continued.</p>

<p>We expect all students to attend all upcoming live training sessions promptly.</p>

<p>We urge you to prioritize your attendance and actively participate in these valuable sessions.</p>

<p>Regards,</p>
</div>`;

    const nietStaffEmails = 'premsagar.sharma@niet.co.in, "Dr. Neema Agarwal" <amd@niet.co.in>, "Dr. Vinod M Kapse" <director@niet.co.in>, arvind.sharma@niet.co.in, ajeet.singh@niet.co.in';
    const myAnatomyStaffEmails = 'nishi.s@myanatomy.in, sucharita@myanatomy.in';
    
    let absentStudentEmails = '';
    if (batchStats && batchStats.absentStudents.length > 0) {
      absentStudentEmails = batchStats.absentStudents.map(student => student.email).filter(email => email).join(', ');
    }

    setAbsentStudentEmailContent(htmlContentForDisplay);
    setAbsentStudentEmailContentForCopy(htmlContentForCopy);
    setAbsentStudentEmailSubject(subjectLine);
    setAbsentStudentEmailTo(nietStaffEmails);
    setAbsentStudentEmailCC(myAnatomyStaffEmails);
    setAbsentStudentEmailBCC(absentStudentEmails);
  }, [selectedEmailBatchSheet, internReport, allSheetsAttendanceData]);;

  const formatInternReportToTable = (content: string): string => {
    if (!content.trim()) return '';
    
    const lines = content.split('\n').filter(line => line.trim());
    
    const getTopicDisplayName = (sheetName: string): string => {
      const sheetLower = sheetName.toLowerCase();
      if (sheetLower.includes('ms-1') || sheetLower.includes('mern')) {
        return 'MERN Stack';
      } else if (sheetLower.includes('java sde-1') || sheetLower.includes('java sde 1')) {
        return 'JAVA SDE';
      } else if (sheetLower.includes('java sde-2') || sheetLower.includes('java sde 2')) {
        return 'JAVA SDE';
      } else if (sheetLower.includes('data scientist') || sheetLower.includes('python')) {
        return 'Data Science';
      }
      return sheetName; // Fallback to original name if no match
    };

    const topicDisplayName = getTopicDisplayName(selectedEmailBatchSheet);

    // Create table header
    let tableHTML = `
    <table style="width: 100%; border-collapse: collapse; margin: 10px 0;">
      <thead>
        <tr style="background-color: #FFE100;">
          <th style="border: 1px solid #000000; padding: 8px; text-align: center; font-weight: bold;">S.No</th>
          <th style="border: 1px solid #000000; padding: 8px; text-align: center; font-weight: bold;">Topic</th>
          <th style="border: 1px solid #000000; padding: 8px; text-align: left; font-weight: bold;">Description</th>
        </tr>
      </thead>
      <tbody>`;
    
    // Add table row - only for the primary selected batch
    if (selectedSheet) {
      const numberedDescription = lines.map((line, lineIndex) => `${lineIndex + 1}. ${line.trim()}`).join('<br>');
      tableHTML += `
        <tr style="background-color: #ffffff;">
          <td style="border: 1px solid #000000; padding: 8px; text-align: center;">1</td>
          <td style="border: 1px solid #000000; padding: 8px; text-align: center; font-weight: bold;">${topicDisplayName}</td>
          <td style="border: 1px solid #000000; padding: 8px; text-align: left;">${numberedDescription}</td>
        </tr>`;
    }
    
    tableHTML += `
      </tbody>
    </table>`;
    
    return tableHTML;
  };

  const generatePresentStudentEmail = useCallback(() => {
    if (!selectedEmailBatchSheet || !allSheetsAttendanceData.has(selectedEmailBatchSheet)) return;

    const batchStats = allSheetsAttendanceData.get(selectedEmailBatchSheet);
    if (!batchStats) return; // Should not happen due to has() check, but for type safety
    
    const formatDateForEmail = (dateStr: string): string => {
      const [day, month, year] = dateStr.split('-');
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const monthName = monthNames[parseInt(month) - 1];
      return `${day} ${monthName} ${year}`;
    };

    const formattedDate = formatDateForEmail(batchStats.date);
    
    // Map sheet names to proper topic names for subject line
    const getTopicName = (sheetName: string): string => {
      const name = sheetName.toLowerCase();
      if (name.includes('ms-1') || name.includes('mern')) {
        return 'MERN Stack';
      } else if (name.includes('java sde-1') || name.includes('java sde 1')) {
        return 'JAVA SDE-1';
      } else if (name.includes('java sde-2') || name.includes('java sde 2')) {
        return 'JAVA SDE-2';
      } else if (name.includes('data scientist') || name.includes('python')) {
        return 'Data Science Python';
      }
      return sheetName; // fallback to original name
    };
    
    const topicName = getTopicName(selectedEmailBatchSheet);
    const subjectLine = `${topicName} NCET + Online Training NIET College Attendance ${formattedDate}`;
    
    // Version for display (light text for dark UI)
    const htmlContentForDisplay = `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #E5E7EB;">
<p><strong>Dear Students</strong>,</p>

<p>On behalf of the <strong>NCET Live Training</strong> team, we would like to congratulate you on your punctuality in attending the recent live training session conducted on <strong>${formattedDate}</strong>.</p>

<p>We appreciate your dedication and commitment to learning.</p>

<p>Here's a quick recap of what was discussed during the session:</p>

${internReport ? formatInternReportToTable(internReport) : '<p><em>Session summary will be added here based on the intern report content.</em></p>'}

<p>We have also received feedback from many of you regarding the sessions and go through it continuously to identify how we can improve the process.</p>

<p>Thank you for your continued support and participation.</p>

<p>Regards,</p>
</div>`;

    // Version for copying (black text for email clients)
    const htmlContentForCopy = `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #000000;">
<p><strong>Dear Students</strong>,</p>

<p>On behalf of the <strong>NCET Live Training</strong> team, we would like to congratulate you on your punctuality in attending the recent live training session conducted on <strong>${formattedDate}</strong>.</p>


<p>We appreciate your dedication and commitment to learning.</p>

<p>Here's a quick recap of what was discussed during the session:</p>

${internReport ? formatInternReportToTable(internReport) : '<p><em>Session summary will be added here based on the intern report content.</em></p>'}

<p>We have also received feedback from many of you regarding the sessions and go through it continuously to identify how we can improve the process.</p>

<p>Thank you for your continued support and participation.</p>

<p>Regards,</p>
</div>`;

    const nietStaffEmails = 'premsagar.sharma@niet.co.in, "Dr. Neema Agarwal" <amd@niet.co.in>, "Dr. Vinod M Kapse" <director@niet.co.in>, arvind.sharma@niet.co.in, ajeet.singh@niet.co.in';
    const myAnatomyStaffEmails = 'nishi.s@myanatomy.in, sucharita@myanatomy.in';
    
    let presentStudentEmails = '';
    if (batchStats && batchStats.presentStudents.length > 0) {
      presentStudentEmails = batchStats.presentStudents.map(student => student.email).filter(email => email).join(', ');
    }

    setPresentStudentEmailContent(htmlContentForDisplay);
    setPresentStudentEmailContentForCopy(htmlContentForCopy);
    setPresentStudentEmailSubject(subjectLine);
    setPresentStudentEmailTo(nietStaffEmails);
    setPresentStudentEmailCC(myAnatomyStaffEmails);
    setPresentStudentEmailBCC(presentStudentEmails);
  }, [selectedEmailBatchSheet, internReport, allSheetsAttendanceData]);

  // Copy functions
  const copyAbsentStudentEmail = async () => {
    if (!absentStudentEmailContentForCopy) {
      generateAbsentStudentEmail();
      return;
    }

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([absentStudentEmailContentForCopy], { type: 'text/html' }),
          'text/plain': new Blob([absentStudentEmailContentForCopy.replace(/<[^>]*>/g, '')], { type: 'text/plain' })
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
    if (!presentStudentEmailContentForCopy) {
      generatePresentStudentEmail();
      return;
    }

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([presentStudentEmailContentForCopy], { type: 'text/html' }),
          'text/plain': new Blob([presentStudentEmailContentForCopy.replace(/<[^>]*>/g, '')], { type: 'text/plain' })
        })
      ];
      await navigator.clipboard.write(clipboardData);
      setCopiedPresentStudentEmail(true);
      setTimeout(() => setCopiedPresentStudentEmail(false), 2000);
    } catch (error) {
      console.error('Failed to copy present student email:', error);
    }
  };

  // Get students from selected batch only
  const getSelectedBatchStudents = (present: boolean) => {
    if (!selectedEmailBatchSheet) {
      // If no batch is selected for email, return empty list
      return [];
    }
    
    const batchStats = allSheetsAttendanceData.get(selectedEmailBatchSheet);
    
    if (!batchStats) {
      // If stats for the selected batch are not available, return empty list
      return [];
    }
    
    return present ? batchStats.presentStudents : batchStats.absentStudents;
  };

  const copyAbsentEmails = async () => {
    const students = getSelectedBatchStudents(false);
    if (students.length > 0) {
      const emails = students.map(student => student.email).join(', ');
      try {
        await navigator.clipboard.writeText(emails);
        setCopiedAbsentEmails(true);
        setTimeout(() => setCopiedAbsentEmails(false), 2000);
      } catch (error) {
        console.error('Failed to copy absent emails:', error);
      }
    }
  };

  const copyPresentEmails = async () => {
    const students = getSelectedBatchStudents(true);
    if (students.length > 0) {
      const emails = students.map(student => student.email).join(', ');
      try {
        await navigator.clipboard.writeText(emails);
        setCopiedPresentEmails(true);
        setTimeout(() => setCopiedPresentEmails(false), 2000);
      } catch (error) {
        console.error('Failed to copy present emails:', error);
      }
    }
  };

  const generateEmailTemplate = () => {
    const actualBatches = generateAllBatchDataFromAttendance();
    if (actualBatches.length === 0) return;

    const formatDate = (dateStr: string): string => {
      if (!dateStr) return '';
      const parts = dateStr.split(/[-/]/);
      if (parts.length === 3) {
        return parts.join('/');
      }
      return dateStr;
    };

    const actualDate = actualBatches[0]?.date || selectedDate || '19/08/2025';
    const formattedDate = formatDate(actualDate);

    // Generate batch content similar to your example
    let batchContent = '';
    if (actualBatches.length > 0) {
      batchContent = actualBatches.map(batch => {
        const presentFormatted = String(batch.present).padStart(2, '0');
        const absentFormatted = String(batch.absent).padStart(2, '0');
        return `<strong>·        Total Number of Registered Students for ${batch.name}: ${batch.total}<br>·        Number of Students Present for ${batch.name}: ${presentFormatted}<br>·        Number of Students Absent for ${batch.name}: ${absentFormatted}</strong>`;
      }).join('<br><br>');
    }

    const content = `<p>Dear Ma'am/Sir,<br>Greetings of the day!<br>I hope you are doing well.</p><p>This is to inform you that the training session conducted on ${formattedDate} for the <strong>NCET + Training</strong> was successfully completed. Please find below the attendance details of the students who participated in the session:</p><p><strong>${batchContent}</strong></p><p>I’m sharing the sheet attached in the email for your reference. It includes <strong>detailed attendance, daily Assessment, and a list of absent students</strong> during each training session.</p><p><strong>Link for Daily Attendance and Assessment:</strong> ${emailTemplate.sheetsLink}</p><p>Kindly go through the same and let us know if you have any questions or need any further information.</p><p>Thank you for your continued support and coordination.</p><p><strong>Regards,</strong></p>`;

    setEmailTemplate(prev => ({ 
      ...prev, 
      trainingDate: formattedDate,
      batches: actualBatches.map(b => b.name),
      generatedContent: content 
    }));
    
    // Also set the subject line separately
    setEmailTemplateSubject(`NIET College NCET + Training Attendance for ${formattedDate}`);
  };

  const copyEmailTemplate = async () => {
    if (!emailTemplate.generatedContent) return;

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([emailTemplate.generatedContent], { type: 'text/html' }),
          'text/plain': new Blob([emailTemplate.generatedContent.replace(/<[^>]*>/g, '')], { type: 'text/plain' })
        })
      ];
      await navigator.clipboard.write(clipboardData);
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

  const copyEmailCc = async () => {
    try {
      await navigator.clipboard.writeText(emailTemplate.cc);
      setCopiedEmailCc(true);
      setTimeout(() => setCopiedEmailCc(false), 2000);
    } catch (error) {
      console.error('Failed to copy email cc:', error);
    }
  };

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

  const copyAbsentStudentSubject = async () => {
    if (!absentStudentEmailSubject) return;
    try {
      await navigator.clipboard.writeText(absentStudentEmailSubject);
      setCopiedAbsentStudentSubject(true);
      setTimeout(() => setCopiedAbsentStudentSubject(false), 2000);
    } catch (error) {
      console.error('Failed to copy absent student subject:', error);
    }
  };

  const copyPresentStudentSubject = async () => {
    if (!presentStudentEmailSubject) return;
    try {
      await navigator.clipboard.writeText(presentStudentEmailSubject);
      setCopiedPresentStudentSubject(true);
      setTimeout(() => setCopiedPresentStudentSubject(false), 2000);
    } catch (error) {
      console.error('Failed to copy present student subject:', error);
    }
  };

  const copyAbsentStudentTo = async () => {
    if (!absentStudentEmailTo) return;
    await navigator.clipboard.writeText(absentStudentEmailTo);
    setCopiedAbsentStudentTo(true);
    setTimeout(() => setCopiedAbsentStudentTo(false), 2000);
  };

  const copyAbsentStudentCc = async () => {
    if (!absentStudentEmailCC) return;
    await navigator.clipboard.writeText(absentStudentEmailCC);
    setCopiedAbsentStudentCc(true);
    setTimeout(() => setCopiedAbsentStudentCc(false), 2000);
  };

  const copyAbsentStudentBcc = async () => {
    if (!absentStudentEmailBCC) return;
    await navigator.clipboard.writeText(absentStudentEmailBCC);
    setCopiedAbsentStudentBcc(true);
    setTimeout(() => setCopiedAbsentStudentBcc(false), 2000);
  };

  const copyPresentStudentTo = async () => {
    if (!presentStudentEmailTo) return;
    await navigator.clipboard.writeText(presentStudentEmailTo);
    setCopiedPresentStudentTo(true);
    setTimeout(() => setCopiedPresentStudentTo(false), 2000);
  };

  const copyPresentStudentCc = async () => {
    if (!presentStudentEmailCC) return;
    await navigator.clipboard.writeText(presentStudentEmailCC);
    setCopiedPresentStudentCc(true);
    setTimeout(() => setCopiedPresentStudentCc(false), 2000);
  };

  const copyPresentStudentBcc = async () => {
    if (!presentStudentEmailBCC) return;
    await navigator.clipboard.writeText(presentStudentEmailBCC);
    setCopiedPresentStudentBcc(true);
    setTimeout(() => setCopiedPresentStudentBcc(false), 2000);
  };

  const generateAllBatchDataFromAttendance = () => {
    const batchData: Array<{
      id: number;
      name: string;
      present: number;
      absent: number;
      total: number;
      percentage: string;
      date: string;
    }> = [];
    let batchId = 1;
    
    for (const [sheetName, sheetStats] of allSheetsAttendanceData.entries()) {
      let batchName = sheetName; // Default to sheetName
      const sheetLower = sheetName.toLowerCase();
      
      if (sheetLower.includes('ms1 attendance') || sheetLower.includes('ms-1 attendance')) {
        batchName = 'MERN Stack Batch 1';
      } else if (sheetLower.includes('java sde-1 attendance')) {
        batchName = 'Java SDE-1 Batch 2';
      } else if (sheetLower.includes('java sde-2 attendance')) {
        batchName = 'Java SDE-2 Batch 3';
      } else if (sheetLower.includes('data scientist attendance')) {
        batchName = 'Data Science Python Batch 4';
      }
      
      batchData.push({
        id: batchId++,
        name: batchName,
        present: sheetStats.present,
        absent: sheetStats.absent,
        total: sheetStats.totalStudents,
        percentage: sheetStats.presentPercentage.toFixed(1),
        date: sheetStats.date
      });
    }
    
    return batchData;
  };


  return (
    <div className="space-y-8">
      {/* Row 1: Upload and Student Selection - Same structure as NIET */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Box 1: Upload Document Section */}
        <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
          <div className="flex items-center gap-3 mb-4">
            <Upload className="w-5 h-5 text-gray-400" />
            <h2 className="text-lg font-semibold text-white">Upload Files</h2>
            <span className="bg-blue-600 text-white text-xs px-2 py-1 rounded-full">
              NIET
            </span>
          </div>
          
          <label htmlFor="attendance-sheet" className="text-sm text-gray-400 mb-2 block">Attendance Sheet</label>
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
              <div className="text-center text-blue-400">
                <div className="inline-block animate-spin rounded-full h-6 w-6 border-b-2 border-blue-400 mr-2"></div>
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

          {/* Batch Selection - Similar to NIET */}
          {availableSheets.length > 0 && (
            <div className="mt-4 p-4 bg-blue-600/10 border border-blue-500/30 rounded-lg">
              <div className="flex items-center justify-between mb-3">
                <div>
                  <h3 className="text-sm font-semibold text-blue-300 mb-1">Batch Selection</h3>
                  <p className="text-xs text-blue-200/80">
                    {availableSheets.length > 1 ? 'Multiple batches detected. Primary batch for attendance analysis:' : 'Select batch for attendance analysis:'}
                  </p>
                </div>
                <span className="bg-blue-600 text-white text-xs px-2 py-1 rounded-full">
                  {availableSheets.length} batches
                </span>
              </div>
              
              {/* Primary Sheet Selection */}
              <div className="mb-4">
                <label className="text-xs text-blue-200 font-medium mb-2 block">Primary Batch (for date selection):</label>
                <select
                  value={selectedSheet}
                  onChange={(e) => handleSheetChange(e.target.value)}
                  className="w-full bg-gray-700 border border-gray-600 rounded-md px-3 py-2 text-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 select-auto"
                  style={{ userSelect: 'auto', WebkitUserSelect: 'auto' }}
                >
                  {availableSheets.map((sheetName) => (
                    <option key={sheetName} value={sheetName} className="text-white bg-gray-700">
                      📊 {sheetName}
                    </option>
                  ))}
                </select>
              </div>
              
              <div className="flex items-center justify-between mb-3">
                <h4 className="text-sm font-semibold text-white">Select Batches to Process</h4>
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
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
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
                      <div className="text-xs text-blue-400">
                        NIET Batch
                      </div>
                    </div>
                  </label>
                ))}
              </div>
            </div>
          )}
        </section>

        {/* Box 2: Attendance Analysis - Same structure as NIET */}
        <section className={`bg-gray-800/50 border border-gray-700/50 rounded-xl p-6 transition-opacity ${!isUploadComplete ? 'opacity-50' : 'opacity-100'}`}>
          <div className="flex items-center gap-3 mb-4">
            <Book className="w-5 h-5 text-gray-400" />
            <h2 className="text-lg font-semibold text-white">Attendance Analysis</h2>
          </div>
          <p className="text-sm text-gray-400 mb-4">Select a date to view attendance statistics.</p>
          
          {(() => {
            console.log('UI RENDER DEBUG - Attendance Analysis render check:', {
              isUploadComplete,
              attendanceDatesLength: attendanceDates.length,
              allSheetsDataSize: allSheetsData.size,
              selectedSheetsSize: selectedSheetsForProcessing.size,
              attendanceStats: !!attendanceStats
            });
            return isUploadComplete && attendanceDates.length > 0;
          })() ? (
            <div className="space-y-4">
              <div className="flex justify-between items-center">
                <span className="text-sm text-gray-300">
                  {attendanceDates.length} attendance dates available
                </span>
                <div className="flex gap-2">
                  <span className="text-xs px-2 py-1 bg-blue-600 text-white rounded">
                    NIET
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
                  
                  {/* All Batches Summary */}
                  {allSheetsAttendanceData.size > 0 && (
                    <div className="mt-6 bg-gray-600/30 rounded-lg p-4">
                      <h4 className="text-sm font-semibold text-white mb-3">
                        All Batches Summary for {attendanceStats.date}
                      </h4>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                        {Array.from(allSheetsAttendanceData.entries()).map(([batchName, stats]) => (
                          <div key={batchName} className="bg-gray-700/50 rounded p-3">
                            <div className="text-xs font-medium text-blue-300 mb-2">{batchName}</div>
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
                            Present Students ({getSelectedBatchStudents(true).length})
                            {selectedEmailBatchSheet && (
                              <span className="text-xs text-green-400 ml-2">({selectedEmailBatchSheet})</span>
                            )}
                          </h4>
                          <p className="text-xs text-green-400/70 mt-1">
                            📧 Ready for Gmail &quot;BCC&quot; field
                            {selectedEmailBatchSheet ? ` - ${selectedEmailBatchSheet} only` : ' - All batches'}
                          </p>
                        </div>
                        <button
                          onClick={copyPresentEmails}
                          disabled={getSelectedBatchStudents(true).length === 0}
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
                        {getSelectedBatchStudents(true).length > 0 ? (
                          getSelectedBatchStudents(true).map((student, index) => (
                            <div key={index} className="bg-green-700/20 rounded px-3 py-2">
                              <div className="text-sm text-white font-medium">{student.name}</div>
                              <div className="text-xs text-green-300 font-mono">{student.email}</div>
                            </div>
                          ))
                        ) : (
                          <div className="text-center text-green-400 text-sm py-4">
                            {selectedEmailBatchSheet ? `No students present in ${selectedEmailBatchSheet}` : 'No students present'}
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Absent Students */}
                    <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
                      <div className="flex items-center justify-between mb-3">
                        <div>
                          <h4 className="text-red-300 text-sm font-semibold">
                            Absent Students ({getSelectedBatchStudents(false).length})
                            {selectedEmailBatchSheet && (
                              <span className="text-xs text-red-400 ml-2">({selectedEmailBatchSheet})</span>
                            )}
                          </h4>
                          <p className="text-xs text-red-400/70 mt-1">
                            📧 Ready for Gmail &quot;BCC&quot; field
                            {selectedEmailBatchSheet ? ` - ${selectedEmailBatchSheet} only` : ' - All batches'}
                          </p>
                        </div>
                        <button
                          onClick={copyAbsentEmails}
                          disabled={getSelectedBatchStudents(false).length === 0}
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
                        {getSelectedBatchStudents(false).length > 0 ? (
                          getSelectedBatchStudents(false).map((student, index) => (
                            <div key={index} className="bg-red-700/20 rounded px-3 py-2">
                              <div className="text-sm text-white font-medium">{student.name}</div>
                              <div className="text-xs text-red-300 font-mono">{student.email}</div>
                            </div>
                          ))
                        ) : (
                          <div className="text-center text-red-400 text-sm py-4">
                            {selectedEmailBatchSheet ? `No students absent in ${selectedEmailBatchSheet}` : 'No students absent'}
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

      {/* Row 4: Email Template Generator - Full Width - EXACT NIET Structure */}
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
            {emailTemplate.generatedContent && (
              <button
                onClick={copyEmailTemplate}
                className="flex items-center gap-2 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-sm rounded-md transition-colors"
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
            )}
          </div>
        </div>
        
        <p className="text-sm text-gray-400 mb-6">
          Create a comprehensive email report with attendance statistics, batch data, and professional formatting. Perfect for faculty communication.
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

              {/* To Field */}
              <div className="space-y-2">
                <label className="text-xs text-gray-300 font-medium">To:</label>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    value={emailTemplate.to}
                    onChange={(e) => setEmailTemplate(prev => ({ ...prev, to: e.target.value }))}
                    placeholder="recipient@example.com"
                    className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-orange-500 focus:border-orange-500 select-text"
                    style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                  />
                  <button
                    onClick={copyEmailTo}
                    className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors"
                  >
                    {copiedEmailTo ? (
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

              {/* CC Field */}
              <div className="space-y-2">
                <label className="text-xs text-gray-300 font-medium">CC:</label>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    value={emailTemplate.cc}
                    onChange={(e) => setEmailTemplate(prev => ({ ...prev, cc: e.target.value }))}
                    placeholder="recipient@example.com"
                    className="w-full bg-gray-600 border border-gray-500 rounded px-3 py-2 text-white text-sm focus:ring-2 focus:ring-orange-500 focus:border-orange-500 select-text"
                    style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                  />
                  <button
                    onClick={copyEmailCc}
                    className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors"
                  >
                    {copiedEmailCc ? (
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
                      {generateAllBatchDataFromAttendance().map((batch) => (
                        <div key={batch.id} className="bg-gray-600/50 rounded p-3 text-xs">
                          <div className="flex justify-between items-center mb-1">
                            <span className="text-white font-medium">{batch.name}</span>
                            <span className={`px-2 py-1 rounded text-xs ${
                              parseFloat(batch.percentage) >= 75 
                                ? 'bg-green-600/20 text-green-300' 
                                : parseFloat(batch.percentage) >= 50
                                ? 'bg-yellow-600/20 text-yellow-300'
                                : 'bg-red-600/20 text-red-300'
                            }`}>
                              {batch.percentage}%
                            </span>
                          </div>
                          <div className="text-gray-300 space-y-1">
                            <div>Present: <span className="text-green-300">{batch.present}</span></div>
                            <div>Absent: <span className="text-red-300">{batch.absent}</span></div>
                            <div>Total: <span className="text-blue-300">{batch.total}</span></div>
                          </div>
                        </div>
                      ))}
                    </>
                  ) : (
                    <div className="text-center py-4">
                      <p className="text-xs text-gray-400">No batches selected for processing</p>
                    </div>
                  )}
                </div>
              ) : (
                <div className="text-center py-4">
                  <p className="text-xs text-gray-400">Select a date in attendance analysis to see batch data</p>
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
              
              {/* Email Content */}
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
                <li>Click &quot;Generate Template&quot; to create the email content with formatting</li>
                <li>Use &quot;Copy Email&quot; to copy HTML-formatted content to your clipboard</li>
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

      {/* Row 3: Intern Report Section - Full Width like NIET */}
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
              onChange={(e) => setInternReport(e.target.value)}
              placeholder="Paste your intern report content here...&#10;&#10;You can include:&#10;- Student names and details&#10;- Internship progress&#10;- Performance evaluations&#10;- Any other relevant information"
              className={`w-full bg-gray-700 border border-gray-600 rounded-lg px-4 py-3 text-white placeholder-gray-400 focus:ring-2 focus:ring-purple-500 focus:border-purple-500 resize-y select-text ${internReportExpanded ? 'h-64' : 'h-48'}`}
              style={{ userSelect: 'text', WebkitUserSelect: 'text', minHeight: '12rem' }}
            />
            <div className="flex justify-between items-center mt-2">
              <span className="text-xs text-gray-500">
                {internReport.length} characters, {internReport.split('\n').filter(line => line.trim()).length} lines
              </span>
              <button
                onClick={() => setInternReport('')}
                className="text-xs text-red-400 hover:text-red-300"
              >
                Clear All
              </button>
            </div>
          </div>
          
          {/* Preview Section */}
          {internReport && internReportExpanded && (
            <div className="bg-gray-700/30 rounded-lg p-4">
              <h4 className="text-sm font-semibold text-white mb-3">Processed Entries Preview</h4>
              <div className="max-h-48 overflow-y-auto space-y-2">
                {internReport.split('\n').filter(line => line.trim()).map((line, index) => (
                  <div key={index} className="text-xs text-gray-300 p-2 bg-gray-600/50 rounded">
                    <span className="text-purple-400 font-semibold">#{index + 1}</span> {line.trim()}
                  </div>
                ))}
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
                  <li>Use &quot;Expand&quot; to view the processed entries</li>
                  <li>Each line of meaningful content becomes a separate entry</li>
                </ul>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Row 5: Student Email Templates - Full Width like NIET */}
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
            Choose which batch to include in the BCC list (will show students from selected batch only)
          </p>
          <select
            value={selectedEmailBatchSheet || ''}
            onChange={(e) => setSelectedEmailBatchSheet(e.target.value)}
            className="w-full px-3 py-2 bg-gray-700 border border-gray-600 rounded-md text-white text-sm focus:ring-2 focus:ring-purple-500 focus:border-transparent"
          >
            <option value="">Select a batch...</option>
            {availableSheets.map((sheet) => (
              <option key={sheet} value={sheet}>{sheet}</option>
            ))}
          </select>
        </div>

        <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
          {/* Absent Students Email */}
          <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
            <div className="flex items-center justify-between mb-4">
              <div>
                <div className='flex flex-row'>
                  <h3 className="text-lg font-semibold text-red-300 mb-2">Email for Absent Students</h3>
                  {selectedEmailBatchSheet && allSheetsAttendanceData.has(selectedEmailBatchSheet) && (
                <div className="flex items-center gap-2 ml-4 -mt-1.5">
                  <div className="bg-red-600/20 text-red-300 text-xs px-3 py-1 rounded-full">
                    Absent: {allSheetsAttendanceData.get(selectedEmailBatchSheet)?.absent}
                  </div>
                </div>
                )}
                </div>
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
              <div className="space-y-3">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">Subject:</span>
                  <input type="text" value={absentStudentEmailSubject} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentSubject} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentSubject ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">To:</span>
                  <input type="text" value={absentStudentEmailTo} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentTo} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentTo ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">CC:</span>
                  <input type="text" value={absentStudentEmailCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentCc} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentCc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">BCC:</span>
                  <input type="text" value={absentStudentEmailBCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentBcc} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentBcc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="mt-2 text-xs text-red-400 font-medium">
                  {selectedEmailBatchSheet && allSheetsAttendanceData.has(selectedEmailBatchSheet) ? `${allSheetsAttendanceData.get(selectedEmailBatchSheet)?.absent} absent students will be BCC'd` : 'No attendance data available for selected batch'}
                </div>

                <div className="bg-gray-800 rounded p-3 max-h-64 overflow-y-auto">
                  <div 
                    className="text-xs leading-relaxed"
                    dangerouslySetInnerHTML={{ __html: absentStudentEmailContent }}
                  />
                </div>
              </div>
            )}
          </div>

          {/* Present Students Email */}
          <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
            <div className="flex items-center justify-between mb-4">
              <div>
                <div className='flex flex-row'>
                  <h3 className="text-lg font-semibold text-green-300 mb-2">Email for Present Students</h3>
                  {selectedEmailBatchSheet && allSheetsAttendanceData.has(selectedEmailBatchSheet) && (
                <div className="flex items-center gap-2 ml-4 -mt-2">
                  <div className="bg-green-600/20 text-green-300 text-xs px-3 py-1 rounded-full">
                    Present: {allSheetsAttendanceData.get(selectedEmailBatchSheet)?.present}
                  </div>
                </div>
                )}
                </div>
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
              <div className="space-y-3">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">Subject:</span>
                  <input type="text" value={presentStudentEmailSubject} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentSubject} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentSubject ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">To:</span>
                  <input type="text" value={presentStudentEmailTo} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentTo} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentTo ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">CC:</span>
                  <input type="text" value={presentStudentEmailCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentCc} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentCc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">BCC:</span>
                  <input type="text" value={presentStudentEmailBCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentBcc} className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentBcc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="mt-2 text-xs text-green-400 font-medium">
                  {selectedEmailBatchSheet && allSheetsAttendanceData.has(selectedEmailBatchSheet) ? `${allSheetsAttendanceData.get(selectedEmailBatchSheet)?.present} present students will be BCC'd` : 'No attendance data available for selected batch'}
                </div>

                <div className="bg-gray-800 rounded p-3 max-h-64 overflow-y-auto">
                  <div 
                    className="text-xs leading-relaxed"
                    dangerouslySetInnerHTML={{ __html: presentStudentEmailContent }}
                  />
                </div>
              </div>
            )}
          </div>
        </div>
        
        {/* Help Section like NIET */}
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
    </div>
  );
}