'use client';

import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { FileText, Upload, Book, Copy, Check, Mail, ChevronDown, Users } from 'lucide-react';

interface SRMOfflineStudent {
  serialNo: string;
  regnNumber: string;
  name: string;
  email: string;
  program: string;
  attendance: { [key: string]: number }; // date -> 0/1
}

interface SRMOfflineAttendanceStats {
  date: string;
  totalStudents: number;
  present: number;
  absent: number;
  presentPercentage: number;
  absentPercentage: number;
  presentStudents: Array<{ name: string; email: string; regnNumber: string; program: string }>;
  absentStudents: Array<{ name: string; email: string; regnNumber: string; program: string }>;
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
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [selectedEmailBatchSheet, setSelectedEmailBatchSheet] = useState<string>(''); // For student email templates

  // Email states
  const [absentStudentEmailContent, setAbsentStudentEmailContent] = useState<string>('');
  const [presentStudentEmailContent, setPresentStudentEmailContent] = useState<string>('');

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
    
    const presentStudents: Array<{ name: string; email: string; regnNumber: string; program: string }> = [];
    const absentStudents: Array<{ name: string; email: string; regnNumber: string; program: string }> = [];
    
    allStudents.forEach(student => {
      const attendanceValue = student.attendance[date];
      if (attendanceValue === 1) {
        presentStudents.push({
          name: student.name,
          email: student.email,
          regnNumber: student.regnNumber,
          program: student.program
        });
      } else if (attendanceValue === 0) {
        absentStudents.push({
          name: student.name,
          email: student.email,
          regnNumber: student.regnNumber,
          program: student.program
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
            .map(student => ({ name: student.name, email: student.email, regnNumber: student.regnNumber, program: student.program }));
          const absentStudents = sheetData.filter(student => student.attendance[selectedDate] === 0)
            .map(student => ({ name: student.name, email: student.email, regnNumber: student.regnNumber, program: student.program }));
          
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

  useEffect(() => {
    if (selectedDate) {
      const formattedDate = formatDateForEmail(selectedDate);
      setEmailTemplate(prev => ({ ...prev, trainingDate: formattedDate }));
    }
  }, [selectedDate]);

  if (!isVisible) return null;

  // Functions
  const loadSRMOfflineAttendanceSheet = async () => {
    // Load the pre-configured SRM Offline attendance sheet
    console.log('Loading SRM Offline attendance sheet...');
    
    try {
      setIsProcessing(true);
      setIsUploadComplete(false);
      
      // Fetch the Excel file from public folder
      const response = await fetch('/Attendance & Assement SRM offline 05-08-2025.xlsx');
      if (!response.ok) {
        throw new Error('Failed to load SRM Offline attendance sheet');
      }
      
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      
      console.log('SRM Offline - Available sheet names:', workbook.SheetNames);
      
      // Find the correct sheet name (look for sheet containing "Attendance" or use the first sheet)
      let targetSheetName = 'Attendance';
      if (!workbook.SheetNames.includes('Attendance')) {
        // Look for sheet names that might contain 'Attendance'
        const attendanceSheet = workbook.SheetNames.find(name => 
          name.toLowerCase().includes('attendance') || 
          name.toLowerCase().includes('attend')
        );
        
        if (attendanceSheet) {
          targetSheetName = attendanceSheet;
          console.log(`SRM Offline - Found attendance sheet: ${targetSheetName}`);
        } else {
          // Use the first sheet if no attendance sheet found
          targetSheetName = workbook.SheetNames[0];
          console.log(`SRM Offline - No attendance sheet found, using first sheet: ${targetSheetName}`);
        }
      }
      
      // Process the identified sheet
      processSRMOfflineAttendanceData(workbook, targetSheetName);
      
      setIsUploadComplete(true);
      console.log('SRM Offline attendance sheet loaded successfully');
    } catch (error) {
      console.error('Error loading SRM Offline attendance sheet:', error);
    } finally {
      setIsProcessing(false);
    }
  };

  const processSRMOfflineAttendanceData = (workbook: XLSX.WorkBook, primarySheet: string) => {
    console.log('Processing SRM Offline attendance data from sheet:', primarySheet);
    
    const worksheet = workbook.Sheets[primarySheet];
    if (!worksheet) {
      console.error('Sheet not found:', primarySheet);
      return;
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const students: SRMOfflineStudent[] = [];
    const dateColumns: { [key: number]: string } = {};
    
    console.log('SRM Offline - Total rows in sheet:', jsonData.length);
    console.log('SRM Offline - First row (headers):', jsonData[0]);
    
    if (jsonData.length < 2) {
      console.log('SRM Offline - Sheet has insufficient data');
      return;
    }

    // Get header row (row 1, index 0) to find date columns
    const headerRow = jsonData[0] as unknown[];
    
    // Find date columns starting from column F (index 5)
    for (let i = 5; i < headerRow.length; i++) {
      const cellValue = headerRow[i];
      if (cellValue) {
        const cellStr = String(cellValue).trim();
        console.log(`SRM Offline - Header cell ${i}:`, cellStr);
        
        // Check for DD-MM-YYYY format with hyphens (like "05-08-2025")
        const dateMatch = cellStr.match(/(\d{1,2}[-]\d{1,2}[-]\d{4})/);
        
        // Also check for Excel serial date numbers
        if (!dateMatch && !isNaN(Number(cellStr))) {
          const excelDate = Number(cellStr);
          if (excelDate > 40000 && excelDate < 50000) {
            const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
            const day = String(jsDate.getDate()).padStart(2, '0');
            const month = String(jsDate.getMonth() + 1).padStart(2, '0');
            const year = jsDate.getFullYear();
            dateColumns[i] = `${day}-${month}-${year}`;
            console.log(`SRM Offline - Converted Excel date ${excelDate} to ${day}-${month}-${year}`);
          }
        } else if (dateMatch) {
          dateColumns[i] = dateMatch[1];
          console.log(`SRM Offline - Found date in header: ${dateMatch[1]}`);
        }
      }
    }
    
    console.log('SRM Offline - Found date columns:', dateColumns);
    
    // Process student data from rows 1-132 (indices 0-131)
    let processedCount = 0;
    for (let i = 0; i < Math.min(132, jsonData.length); i++) {
      const row = jsonData[i] as unknown[];
      
      if (!row || row.length < 5) {
        console.log(`SRM Offline - Skipping row ${i + 1}: insufficient columns`);
        continue;
      }
      
      const serialNo = row[0] ? String(row[0]).trim() : '';
      const regnNumber = row[1] ? String(row[1]).trim() : '';
      const name = row[2] ? String(row[2]).trim() : '';
      const email = row[3] ? String(row[3]).trim() : '';
      const program = row[4] ? String(row[4]).trim() : '';
      
      // Skip header rows and empty names
      if (!name || name === '' || name.trim() === '' || 
          name.toLowerCase().includes('name') || 
          name.toLowerCase().includes('s.no') ||
          name.toLowerCase().includes('serial') ||
          name.toLowerCase().includes('student name')) {
        console.log(`SRM Offline - Skipping row ${i + 1}: invalid name "${name}"`);
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
        regnNumber,
        name,
        email,
        program,
        attendance
      });
      
      processedCount++;
      if (processedCount <= 3) {
        console.log(`SRM Offline - Sample student ${processedCount}:`, {
          name,
          email,
          program,
          attendanceDates: Object.keys(attendance).length
        });
      }
    }
    
    console.log(`SRM Offline - Processed ${students.length} students`);
    
    // Set the data for the single "Attendance" sheet
    const allData = new Map<string, SRMOfflineStudent[]>();
    allData.set('Attendance', students);
    setAllSheetsData(allData);
    
    // Extract and set dates (using Map to ensure uniqueness by date string)
    const dateMap = new Map<string, { date: string; fullText: string }>();
    if (students.length > 0) {
      students.slice(0, 10).forEach(student => {
        Object.keys(student.attendance).forEach(date => {
          if (date && date.trim() !== '' && !dateMap.has(date)) {
            dateMap.set(date, { date, fullText: `${date} - Attendance` });
          }
        });
      });
    }
    
    const sortedDates = Array.from(dateMap.values()).sort((a, b) => {
      // Sort dates in descending order (newest first)
      // Convert DD-MM-YYYY to YYYY-MM-DD for proper Date parsing
      const [dayA, monthA, yearA] = a.date.split('-');
      const [dayB, monthB, yearB] = b.date.split('-');
      const dateA = new Date(`${yearA}-${monthA}-${dayA}`);
      const dateB = new Date(`${yearB}-${monthB}-${dayB}`);
      return dateB.getTime() - dateA.getTime();
    });
    
    setAttendanceDates(sortedDates);
    setAvailableSheets(['Attendance']);
    
    if (sortedDates.length > 0) {
      setSelectedDate(sortedDates[0].date);
    }
    
    // Set the "Attendance" sheet as selected for processing
    setSelectedSheetsForProcessing(new Set(['Attendance']));
    
    // Calculate and set attendance stats for the first date if available
    if (sortedDates.length > 0) {
      const stats = calculateAttendanceStats(sortedDates[0].date);
      if (stats) {
        setAttendanceStats(stats);
        setAllSheetsAttendanceData(new Map([['Attendance', stats]]));
      }
    }
    
    console.log('SRM Offline - Processing complete:', {
      studentsCount: students.length,
      datesCount: sortedDates.length,
      selectedDate: sortedDates[0]?.date
    });
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
        
        // For uploaded files, get available sheets and filter to valid batch names
        const allSheetNames = workbook.SheetNames;
        const validSheetNames = ['MS1', 'MS2', 'AI/ML-1', 'AI/ML-2', 'AI/ML1', 'AI/ML2'];
        const sheetNames = allSheetNames.filter(name => validSheetNames.includes(name));
        
        console.log('SRM Offline - All sheet names in workbook:', allSheetNames);
        console.log('SRM Offline - Valid sheet names found:', sheetNames);
        
        setAvailableSheets(sheetNames);
        
        // Process each sheet
        const allData = new Map<string, SRMOfflineStudent[]>();
        const dateMap = new Map<string, { date: string; fullText: string }>();
        
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
                if (date && date.trim() !== '' && !dateMap.has(date)) {
                  dateMap.set(date, { date, fullText: `${date} - ${sheetName}` });
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
        console.log('SRM Offline - Extracted dates:', Array.from(dateMap.values()));
        
        setAllSheetsData(allData);
        const sortedDates = Array.from(dateMap.values()).sort((a, b) => {
          // Sort dates in descending order (newest first)
          // Convert DD-MM-YYYY to YYYY-MM-DD for proper Date parsing
          const [dayA, monthA, yearA] = a.date.split('-');
          const [dayB, monthB, yearB] = b.date.split('-');
          const dateA = new Date(`${yearA}-${monthA}-${dayA}`);
          const dateB = new Date(`${yearB}-${monthB}-${dayB}`);
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

  const formatDateForEmail = (dateStr: string): string => {
    // Handle DD-MM-YYYY format (SRM Offline uses hyphens)
    // First part is day, second part is month, third part is year
    const [day, month, year] = dateStr.split('-');
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[parseInt(month) - 1];
    return `${parseInt(day)} ${monthName} ${year}`;
  };

  const generateEmailTemplate = () => {
    if (!attendanceStats) return;
    
    const { trainingDate, sheetsLink } = emailTemplate;

    const content = `<p>Dear Sir/Ma'am,<br>Greetings of the day!<br>I hope you are doing well.</p><p>This is to inform you that the training session conducted on <strong>${trainingDate}</strong> for <strong>MyAnatomy SRM Offline Training</strong> was successfully completed. Please find below the attendance details of the students who participated in the session:</p><p><strong>· Total Number of Registered Students: ${attendanceStats.totalStudents}<br>· Number of Students Present: ${attendanceStats.present}<br>· Number of Students Absent: ${attendanceStats.absent}</strong></p><p>The detailed attendance sheet and list of absent students is attached with this email for your reference.</p><p><a href="${sheetsLink}">${sheetsLink}</a></p><p>Kindly go through the same and let us know if you have any questions or need any further information.</p><p>Thank you for your continued support and coordination.</p><p>Regards</p>`;

    setEmailTemplate(prev => ({ ...prev, generatedContent: content }));
  };

  const copyEmailTemplate = async () => {
    if (!emailTemplate.generatedContent) return;

    try {
      const clipboardData = [
        new ClipboardItem({
          'text/html': new Blob([emailTemplate.generatedContent], { type: 'text/html' }),
          'text/plain': new Blob([emailTemplate.generatedContent.replace(/<[^>]*>/g, '\n')], { type: 'text/plain' })
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

  // Get students from selected batch only
  const getSelectedBatchStudents = (present: boolean) => {
    if (!attendanceStats || !selectedEmailBatchSheet) {
      return present ? attendanceStats?.presentStudents || [] : attendanceStats?.absentStudents || [];
    }
    
    const sheetData = allSheetsData.get(selectedEmailBatchSheet);
    if (!sheetData) {
      return [];
    }
    
    return sheetData
      .filter(student => {
        const attendanceValue = student.attendance[selectedDate];
        return present ? attendanceValue === 1 : attendanceValue === 0;
      })
      .map(student => ({
        name: student.name,
        email: student.email,
        regnNumber: student.regnNumber,
        program: student.program
      }));
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
      const program = row[4] ? String(row[4]).trim() : '';
      
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
        regnNumber,
        name,
        email,
        program,
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
    const absentStudents = getSelectedBatchStudents(false);
    const emails = absentStudents.map(s => s.email).filter(email => email).join(', ');
    try {
      await navigator.clipboard.writeText(emails);
      setCopiedAbsentEmails(true);
      setTimeout(() => setCopiedAbsentEmails(false), 2000);
    } catch (error) {
      console.error('Failed to copy absent emails:', error);
    }
  };

  const copyPresentEmails = async () => {
    const presentStudents = getSelectedBatchStudents(true);
    const emails = presentStudents.map(s => s.email).filter(email => email).join(', ');
    try {
      await navigator.clipboard.writeText(emails);
      setCopiedPresentEmails(true);
      setTimeout(() => setCopiedPresentEmails(false), 2000);
    } catch (error) {
      console.error('Failed to copy present emails:', error);
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
                  {attendanceDates.map((dateObj, index) => (
                    <option key={`${dateObj.date}-${index}`} value={dateObj.date}>
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
                              <div className="text-xs text-green-400">{student.program}</div>
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
                              <div className="text-xs text-red-400">{student.program}</div>
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

      {/* Row 2: Email Template Generator - Full Width like NMIET */}
      {attendanceStats && (
        <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <Mail className="w-5 h-5 text-gray-400" />
              <h2 className="text-lg font-semibold text-white">Email Template Generator</h2>
              <span className="bg-orange-600 text-white text-xs px-2 py-1 rounded-full">
                Staff Template
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
          
          <p className="text-sm text-gray-400 mb-4">
            Generate professional email templates for attendance summary reports.
          </p>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
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
              </div>
            </div>

            <div className="lg:col-span-2">
              {emailTemplate.generatedContent ? (
                <div className="bg-gray-700/50 rounded-lg p-4">
                  <h3 className="text-sm font-semibold text-white mb-3">Generated Email Template</h3>
                  <div className="bg-gray-800 rounded p-4 max-h-96 overflow-y-auto">
                    <div 
                      className="text-xs text-gray-300 leading-relaxed whitespace-pre-wrap"
                      dangerouslySetInnerHTML={{ __html: emailTemplate.generatedContent }}
                    />
                  </div>
                </div>
              ) : (
                <div className="bg-gray-700/50 rounded-lg p-4 h-96 flex items-center justify-center">
                  <div className="text-center">
                    <Mail className="w-12 h-12 text-gray-600 mx-auto mb-3" />
                    <p className="text-gray-500 text-sm">Click &ldquo;Generate Template&rdquo; to create email content</p>
                  </div>
                </div>
              )}
            </div>
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
        </section>
      )}

      {/* Row 3: Intern Report Section - Full Width like NMIET */}
      {attendanceStats && (
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

          <div className="grid grid-cols-1 gap-6">
            <div className="space-y-4">
              <div className="space-y-2">
                <label htmlFor="intern-report-input" className="text-sm text-gray-300 font-medium mb-2 block">
                  Intern Report Content:
                </label>
                <textarea
                  id="intern-report-input"
                  value={internReport}
                  onChange={(e) => setInternReport(e.target.value)}
                  placeholder="Paste your intern report content here...&#10;&#10;You can include:&#10;- Student names and details&#10;- Internship progress&#10;- Performance evaluations&#10;- Any other relevant information"
                  className={`w-full bg-gray-700 border border-gray-600 rounded-lg px-4 py-3 text-white placeholder-gray-400 focus:ring-2 focus:ring-purple-500 focus:border-purple-500 resize-y select-text ${internReportExpanded ? 'h-64' : 'h-48'}`}
                  style={{ userSelect: 'text', WebkitUserSelect: 'text' }}
                />
                <div className="flex justify-between items-center text-xs text-gray-500">
                  <span>
                    {internReport.length} characters, {internReport.split('\n').filter(line => line.trim()).length} lines
                  </span>
                  <button
                    onClick={() => setInternReport('')}
                    className="text-red-400 hover:text-red-300 transition-colors"
                  >
                    Clear
                  </button>
                </div>
              </div>

              {internReport && internReportExpanded && (
                <div className="bg-gray-700/50 rounded-lg p-4 max-h-48 overflow-y-auto">
                  <h4 className="text-gray-300 text-sm font-medium mb-2">Preview:</h4>
                  <div className="text-xs text-gray-400 whitespace-pre-wrap leading-relaxed select-text" style={{ userSelect: 'text', WebkitUserSelect: 'text' }}>
                    {internReport.split('\n').filter(line => line.trim()).map((line, index) => (
                      <div key={index} className="mb-1">
                        <span className="text-purple-400">{index + 1}.</span> {line.trim()}
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div className="bg-purple-600/10 border border-purple-500/30 rounded-lg p-4">
              <div className="flex items-start gap-3">
                <FileText className="w-5 h-5 text-purple-400 mt-0.5 flex-shrink-0" />
                <div>
                  <h4 className="text-purple-300 text-sm font-semibold mb-2">How to Use Intern Report</h4>
                  <ul className="text-xs text-purple-200/80 space-y-1 list-disc list-inside">
                    <li>Copy your intern report from any document or email</li>
                    <li>Paste it in the text area above - line breaks will be preserved</li>
                    <li>The system automatically numbers each line for better organization</li>
                    <li>Use the expand button to get more writing space if needed</li>
                    <li>Content will be automatically formatted in student email templates</li>
                    <li>Each line becomes a numbered point in the email summary table</li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {/* Row 4: Student Email Templates - Full Width like NMIET */}
      {attendanceStats && (
        <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <Mail className="w-5 h-5 text-gray-400" />
              <h2 className="text-lg font-semibold text-white">Student Email Templates</h2>
              <span className="bg-orange-600 text-white text-xs px-2 py-1 rounded-full">
                For Students
              </span>
            </div>
          </div>
          
          <p className="text-sm text-gray-400 mb-4">
            Generate personalized email templates for students based on their attendance status. Tables are auto-generated from intern report data.
          </p>
          
          {/* Batch Selection for Email Templates */}
          <div className="mb-6 p-4 bg-orange-600/10 border border-orange-500/30 rounded-lg">
            <div className="flex items-center gap-3 mb-3">
              <Users className="w-4 h-4 text-orange-400" />
              <h4 className="text-sm font-semibold text-orange-300">Select Batch for Email Template</h4>
            </div>
            <p className="text-xs text-orange-200/70 mb-3">
              Choose which batch to include in the BCC list (will show students from selected batch only)
            </p>
            <select
              value={selectedEmailBatchSheet || ''}
              onChange={(e) => setSelectedEmailBatchSheet(e.target.value)}
              className="w-full px-3 py-2 bg-gray-700 border border-gray-600 rounded-md text-white text-sm focus:ring-2 focus:ring-orange-500 focus:border-transparent"
            >
              <option value="">Select a batch...</option>
              {availableSheets.map((sheet) => (
                <option key={sheet} value={sheet}>{sheet}</option>
              ))}
            </select>
          </div>

          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <div className='flex flex-row'>
                    <h3 className="text-lg font-semibold text-red-300 mb-2">Email for Absent Students</h3>
                    {attendanceStats && (
                  <div className="flex items-center gap-2 ml-4 -mt-1.5">
                    <div className="bg-red-600/20 text-red-300 text-xs px-3 py-1 rounded-full">
                      Absent: {attendanceStats.absent}
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
                <div className="bg-gray-700/50 rounded p-3 max-h-64 overflow-y-auto">
                  <div 
                    className="text-xs text-gray-300 leading-relaxed"
                    dangerouslySetInnerHTML={{ __html: absentStudentEmailContent }}
                  />
                </div>
              )}
            </div>

            <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <div className='flex flex-row'>
                    <h3 className="text-lg font-semibold text-green-300 mb-2">Email for Present Students</h3>
                    {attendanceStats && (
                  <div className="flex items-center gap-2 ml-4 -mt-1.5">
                    <div className="bg-green-600/20 text-green-300 text-xs px-3 py-1 rounded-full">
                      Present: {attendanceStats.present}
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
                <div className="bg-gray-700/50 rounded p-3 max-h-64 overflow-y-auto">
                  <div 
                    className="text-xs text-gray-300 leading-relaxed"
                    dangerouslySetInnerHTML={{ __html: presentStudentEmailContent }}
                  />
                </div>
              )}
            </div>
          </div>
        </section>
      )}

    </div>
  );
}