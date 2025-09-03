'use client';

import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { FileText, Upload, Book, Copy, Check, Mail, ChevronDown, Users } from 'lucide-react';

interface NMIETStudent {
  serialNo: string;
  name: string;
  email: string;
  course: string;
  branch: string;
  attendance: { [key: string]: number }; // date -> 0/1
}

interface NMIETAttendanceStats {
  date: string;
  totalStudents: number;
  present: number;
  absent: number;
  presentPercentage: number;
  absentPercentage: number;
  presentStudents: Array<{ name: string; email: string; course: string; branch: string }>;
  absentStudents: Array<{ name: string; email: string; course: string; branch: string }>;
}

interface EmailTemplate {
  trainingDate: string;
  batches: string[];
  sheetsLink: string;
  to: string;
  cc: string;
  generatedContent: string;
}

interface NMIETOnlineProps {
  isVisible: boolean;
}

export default function NMIETOnline({ isVisible }: NMIETOnlineProps) {
  // File and data states
  const [attendanceFile, setAttendanceFile] = useState<File | null>(null);
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [selectedSheetsForProcessing, setSelectedSheetsForProcessing] = useState<Set<string>>(new Set());
  const [allSheetsData, setAllSheetsData] = useState<Map<string, NMIETStudent[]>>(new Map());
  const [allSheetsAttendanceData, setAllSheetsAttendanceData] = useState<Map<string, NMIETAttendanceStats>>(new Map());
  const [attendanceDates, setAttendanceDates] = useState<Array<{ date: string; fullText: string }>>([]);
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [attendanceStats, setAttendanceStats] = useState<NMIETAttendanceStats | null>(null);
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

  // Email template states
  const [emailTemplate, setEmailTemplate] = useState<EmailTemplate>({
    trainingDate: '',
    batches: [],
    sheetsLink: 'https://docs.google.com/spreadsheets/d/1O4DaPCXJlupvX9M3rMnOx7GozoGoGX9ICUCo4cUrG9A/edit?gid=0#gid=0',
    to: 'Pragyan Srichandan <pragyan.srichandan@nmiet.ac.in>, sk.safaruddin@nmiet.ac.in, richa.parida@nmiet.ac.in',
    cc: 'Nishi Sharma <nishi.s@myanatomy.in>, Sucharita Mahapatra <sucharita@myanatomy.in>, CHINMAY KUMAR <ckd@myanatomy.in>',
    generatedContent: ''
  });
  const [emailTemplateSubject, setEmailTemplateSubject] = useState<string>('');

  // Intern report states
  const [internReport, setInternReport] = useState<string>('');
  const [internReportExpanded, setInternReportExpanded] = useState<boolean>(false);

  // Copy states
  const [copiedEmailTemplate, setCopiedEmailTemplate] = useState<boolean>(false);
  
  const [copiedEmailTo, setCopiedEmailTo] = useState<boolean>(false);
  const [copiedEmailCc, setCopiedEmailCc] = useState<boolean>(false);
  const [copiedEmailSubject, setCopiedEmailSubject] = useState<boolean>(false);
  const [copiedAbsentStudentSubject, setCopiedAbsentStudentSubject] = useState<boolean>(false);
  const [copiedAbsentStudentTo, setCopiedAbsentStudentTo] = useState<boolean>(false);
  const [copiedAbsentStudentCc, setCopiedAbsentStudentCc] = useState<boolean>(false);
  const [copiedAbsentStudentBcc, setCopiedAbsentStudentBcc] = useState<boolean>(false);
  const [copiedPresentStudentSubject, setCopiedPresentStudentSubject] = useState<boolean>(false);
  const [copiedPresentStudentTo, setCopiedPresentStudentTo] = useState<boolean>(false);
  const [copiedPresentStudentCc, setCopiedPresentStudentCc] = useState<boolean>(false);
  const [copiedPresentStudentBcc, setCopiedPresentStudentBcc] = useState<boolean>(false);

  const loadNMIETAttendanceSheet = async () => {
    console.log('NMIET DEBUG: Loading NMIET attendance sheet...');
    setIsProcessing(true);
    try {
      const response = await fetch('/Attendance_NMIET College_ MERN StackTraining_26 August Onwards.xlsx');
      console.log('NMIET DEBUG: Fetch response status:', response.status);
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const arrayBuffer = await response.arrayBuffer();
      console.log('NMIET DEBUG: File loaded, size:', arrayBuffer.byteLength, 'bytes');
      
      const blob = new Blob([arrayBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const file = new File([blob], 'Attendance_NMIET College_ MERN StackTraining_26 August Onwards.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      setAttendanceFile(file);
      processExcelFile(file);
    } catch (error) {
      console.error('NMIET DEBUG: Failed to load NMIET attendance sheet:', error);
      setIsProcessing(false);
    }
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files[0]) {
      const file = event.target.files[0];
      setAttendanceFile(file);
      setIsProcessing(true);
      processExcelFile(file);
    }
  };

  const processExcelFile = (file: File, sheetNameToUse?: string) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      const workbook = XLSX.read(data, { type: 'array' });
      
      // Extract all sheet names
      const sheetNames = workbook.SheetNames;
      setAvailableSheets(sheetNames);
      
      // For NMIET, always use "Attendance" sheet or the first sheet available
      const sheetName = sheetNameToUse || sheetNames.find(name => name.toLowerCase().includes('attendance')) || sheetNames[0];
      setSelectedSheet(sheetName);
      
      // Process attendance data for NMIET
      processNMIETAttendanceData(workbook, sheetName);
    };
    reader.readAsArrayBuffer(file);
  };

  const processNMIETAttendanceData = (workbook: XLSX.WorkBook, primarySheet: string) => {
    console.log('NMIET DEBUG: Processing attendance data for sheet:', primarySheet);
    // Process the primary sheet to get dates and data
    const primaryWorksheet = workbook.Sheets[primarySheet];
    const primaryRawData = XLSX.utils.sheet_to_json(primaryWorksheet, { header: 1, defval: '' }) as (string | number)[][];
    
    console.log('NMIET DEBUG: Raw data rows:', primaryRawData.length);
    console.log('NMIET DEBUG: First 6 rows:', primaryRawData.slice(0, 6));
    
    // Extract dates from row 4 (index 3), starting from column H (index 7) for NMIET
    const headerRow = primaryRawData[3] as string[];
    console.log('NMIET DEBUG: Header row (index 3):', headerRow);
    const dates: Array<{ date: string; fullText: string }> = [];
    
    for (let i = 7; i < headerRow.length; i++) {
      const cellValue = headerRow[i];
      if (cellValue && typeof cellValue === 'string' && cellValue.trim()) {
        // Extract date from format like "01/09/2025(12 PM to 2 PM)" or "DD/MM/YYYY(time)"
        const dateMatch = cellValue.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
        if (dateMatch) {
          dates.push({
            date: dateMatch[1],
            fullText: cellValue.trim()
          });
        }
      }
    }
    
    console.log('NMIET DEBUG: Dates found:', dates);
    setAttendanceDates(dates);
    
    // Process students data from row 5 to row 95
    const students: NMIETStudent[] = [];
    console.log('NMIET DEBUG: Processing students from row 5 to', Math.min(primaryRawData.length, 95));
    for (let i = 4; i < Math.min(primaryRawData.length, 95); i++) { // Row 5 to 95
      const row = primaryRawData[i];
      if (row && row[1] && row[2]) { // Must have name and email
        const student: NMIETStudent = {
          serialNo: row[0] ? String(row[0]) : `${i}`,
          name: String(row[1]).trim(),
          email: String(row[2]).trim(),
          course: row[4] ? String(row[4]).trim() : '', // Column E
          branch: row[5] ? String(row[5]).trim() : '', // Column F
          attendance: {}
        };
        
        // Extract attendance data for each date
        dates.forEach((dateInfo, dateIndex) => {
          const columnIndex = 7 + dateIndex; // Starting from column H
          const attendanceValue = row[columnIndex];
          // Convert to number and check for 1 (present) or 0 (absent)
          const numValue = Number(attendanceValue);
          if (numValue === 1) {
            student.attendance[dateInfo.date] = 1; // Present
          } else if (numValue === 0) {
            student.attendance[dateInfo.date] = 0; // Absent
          }
          // Log first student's attendance for debugging
          if (i === 4) {
            console.log(`NMIET DEBUG: Student "${student.name}" date "${dateInfo.date}" value:`, attendanceValue, '-> numValue:', numValue);
          }
        });
        
        students.push(student);
      }
    }
    
    console.log('NMIET DEBUG: Students processed:', students.length);
    console.log('NMIET DEBUG: First 3 students:', students.slice(0, 3));
    
    // Store all students data
    setAllSheetsData(new Map([['Attendance', students]]));
    
    // Set the "Attendance" sheet as selected for processing
    setSelectedSheetsForProcessing(new Set(['Attendance']));
    
    // If there are dates, select the first one by default
    if (dates.length > 0) {
      console.log('NMIET DEBUG: Setting first date:', dates[0].date);
      setSelectedDate(dates[0].date);
      const stats = calculateAttendanceStats(dates[0].date, students);
      if (stats) {
        setAttendanceStats(stats);
        setAllSheetsAttendanceData(new Map([['Attendance', stats]]));
      }
    }
    
    console.log('NMIET DEBUG: Upload complete, setting flags');
    setIsUploadComplete(true);
    setIsProcessing(false);
  };

  const calculateAttendanceStats = (date: string, students: NMIETStudent[]): NMIETAttendanceStats | null => {
    if (!date || students.length === 0) return null;
    
    let present = 0;
    let absent = 0;
    const presentStudents: Array<{ name: string; email: string; course: string; branch: string }> = [];
    const absentStudents: Array<{ name: string; email: string; course: string; branch: string }> = [];
    
    students.forEach(student => {
      const attendance = student.attendance[date];
      if (attendance === 1) {
        present++;
        presentStudents.push({
          name: student.name,
          email: student.email,
          course: student.course,
          branch: student.branch
        });
      } else if (attendance === 0) {
        absent++;
        absentStudents.push({
          name: student.name,
          email: student.email,
          course: student.course,
          branch: student.branch
        });
      }
    });
    
    const totalStudents = present + absent;
    const stats: NMIETAttendanceStats = {
      date,
      totalStudents,
      present,
      absent,
      presentPercentage: totalStudents > 0 ? Math.round((present / totalStudents) * 100) : 0,
      absentPercentage: totalStudents > 0 ? Math.round((absent / totalStudents) * 100) : 0,
      presentStudents,
      absentStudents
    };
    
    console.log('NMIET DEBUG: Calculated stats for', date, ':', stats);
    return stats;
  };

  // Format date for display
  const formatDateForEmail = (dateStr: string): string => {
    const [day, month, year] = dateStr.split('/');
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[parseInt(month) - 1];
    return `${day} ${monthName} ${year}`;
  };

  const generateAbsentStudentEmail = () => {
    if (!attendanceStats) return;
    
    const formattedDate = formatDateForEmail(attendanceStats.date);
    const subjectLine = `MERN Stack NCET + Online Training NMIET College Attendance ${formattedDate}`;
    
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

    const nmietStaffEmails = 'Pragyan Srichandan <pragyan.srichandan@nmiet.ac.in>, sk.safaruddin@nmiet.ac.in, richa.parida@nmiet.ac.in';
    const myAnatomyStaffEmails = 'Nishi Sharma <nishi.s@myanatomy.in>, Sucharita Mahapatra <sucharita@myanatomy.in>, CHINMAY KUMAR <ckd@myanatomy.in>';
    
    let absentStudentEmails = '';
    if (attendanceStats && attendanceStats.absentStudents.length > 0) {
      absentStudentEmails = attendanceStats.absentStudents.map(student => student.email).filter(email => email).join(', ');
    }

    setAbsentStudentEmailContent(htmlContentForDisplay);
    setAbsentStudentEmailContentForCopy(htmlContentForCopy);
    setAbsentStudentEmailSubject(subjectLine);
    setAbsentStudentEmailTo(nmietStaffEmails);
    setAbsentStudentEmailCC(myAnatomyStaffEmails);
    setAbsentStudentEmailBCC(absentStudentEmails);
  };

  const formatInternReportToTable = (content: string): string => {
    if (!content.trim()) return '';
    
    const lines = content.split('\n').filter(line => line.trim());
    
    // Create table header
    let tableHTML = `
    <table style="width: 100%; border-collapse: collapse; margin: 10px 0;">
      <thead>
        <tr style="background-color: #FFE100;">
          <th style="border: 1px solid #000000; padding: 8px; text-align: center; font-weight: bold; color: #000000;">S.No</th>
          <th style="border: 1px solid #000000; padding: 8px; text-align: center; font-weight: bold; color: #000000;">Topic</th>
          <th style="border: 1px solid #000000; padding: 8px; text-align: left; font-weight: bold; color: #000000;">Description</th>
        </tr>
      </thead>
      <tbody>`;
    
    // Add table row - only for the selected sheet
    if (selectedSheet) {
      const numberedDescription = lines.map((line, lineIndex) => `${lineIndex + 1}. ${line.trim()}`).join('<br>');
      tableHTML += `
        <tr style="background-color: #ffffff;">
          <td style="border: 1px solid #000000; padding: 8px; text-align: center; color: #000000;">1</td>
          <td style="border: 1px solid #000000; padding: 8px; text-align: center; font-weight: bold; color: #000000;">MERN Stack</td>
          <td style="border: 1px solid #000000; padding: 8px; text-align: left; color: #000000;">${numberedDescription}</td>
        </tr>`;
    }
    
    tableHTML += `
      </tbody>
    </table>`;
    
    return tableHTML;
  };

  const generatePresentStudentEmail = () => {
    if (!attendanceStats) return;
    
    const formattedDate = formatDateForEmail(attendanceStats.date);
    const subjectLine = `MERN Stack NCET + Online Training NMIET College Attendance ${formattedDate}`;
    
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

    const nmietStaffEmails = 'Pragyan Srichandan <pragyan.srichandan@nmiet.ac.in>, sk.safaruddin@nmiet.ac.in, richa.parida@nmiet.ac.in';
    const myAnatomyStaffEmails = 'Nishi Sharma <nishi.s@myanatomy.in>, Sucharita Mahapatra <sucharita@myanatomy.in>, CHINMAY KUMAR <ckd@myanatomy.in>';
    
    let presentStudentEmails = '';
    if (attendanceStats && attendanceStats.presentStudents.length > 0) {
      presentStudentEmails = attendanceStats.presentStudents.map(student => student.email).filter(email => email).join(', ');
    }

    setPresentStudentEmailContent(htmlContentForDisplay);
    setPresentStudentEmailContentForCopy(htmlContentForCopy);
    setPresentStudentEmailSubject(subjectLine);
    setPresentStudentEmailTo(nmietStaffEmails);
    setPresentStudentEmailCC(myAnatomyStaffEmails);
    setPresentStudentEmailBCC(presentStudentEmails);
  };

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

  // Copy helper functions for email fields
  const copyAbsentStudentSubject = async () => {
    if (!absentStudentEmailSubject) return;
    await navigator.clipboard.writeText(absentStudentEmailSubject);
    setCopiedAbsentStudentSubject(true);
    setTimeout(() => setCopiedAbsentStudentSubject(false), 2000);
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

  const copyPresentStudentSubject = async () => {
    if (!presentStudentEmailSubject) return;
    await navigator.clipboard.writeText(presentStudentEmailSubject);
    setCopiedPresentStudentSubject(true);
    setTimeout(() => setCopiedPresentStudentSubject(false), 2000);
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

  const generateEmailTemplate = () => {
    if (!attendanceStats) return;

    const { trainingDate, sheetsLink } = emailTemplate;

    const content = `<p>Dear Sir/Ma’am,<br>Greetings of the day!<br>I hope you are doing well.</p><p>This is to inform you that the training session conducted on <strong>${trainingDate}</strong> for <strong>MERN Stack NCET+ Training</strong> was successfully completed. Please find below the attendance details of the students who participated in the session:</p><p><strong>· Total Number of Registered Students: ${attendanceStats.totalStudents}<br>· Number of Students Present: ${attendanceStats.present}<br>· Number of Students Absent: ${attendanceStats.absent}</strong></p><p>The detailed attendance sheet and list of absent students is attached with this email for your reference.</p><p><a href="${sheetsLink}">${sheetsLink}</a></p><p>Kindly go through the same and let us know if you have any questions or need any further information.</p><p>Thank you for your continued support and coordination.</p><p>Regards</p>`;

    setEmailTemplate(prev => ({ ...prev, generatedContent: content }));
    setEmailTemplateSubject(`MERN Stack NCET + Training NMIET College Attendance ${trainingDate}`);
  };

  useEffect(() => {
    if (selectedDate) {
      const formattedDate = formatDateForEmail(selectedDate);
      setEmailTemplate(prev => ({ ...prev, trainingDate: formattedDate }));
    }
  }, [selectedDate]);

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

  if (!isVisible) return null;

  return (
    <div className="space-y-8">
      {/* Row 1: Upload and Student Selection - Same structure as SRM Online */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Box 1: Upload Document Section */}
        <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
          <div className="flex items-center gap-3 mb-4">
            <Upload className="w-5 h-5 text-gray-400" />
            <h2 className="text-lg font-semibold text-white">Upload Files</h2>
            <span className="bg-orange-600 text-white text-xs px-2 py-1 rounded-full">
              NMIET
            </span>
          </div>
          
          <label htmlFor="attendance-sheet" className="text-sm text-gray-400 mb-2 block">Attendance Sheet</label>
          <div className="space-y-4">
            <div className="relative border-2 border-dashed border-orange-600 rounded-lg p-8 text-center bg-orange-600/10">
              <div className="flex flex-col items-center justify-center">
                <FileText className="w-10 h-10 text-orange-500 mb-3" />
                <p className="text-white font-medium mb-2">NMIET Attendance Sheet</p>
                <p className="text-xs text-gray-400 mb-4">
                  {attendanceFile ? `Loaded: ${attendanceFile.name}` : 'Use the pre-configured attendance sheet or upload a custom one'}
                </p>
                <button 
                  onClick={loadNMIETAttendanceSheet}
                  className="px-4 py-2 bg-orange-600 text-white rounded-md hover:bg-orange-700 transition-colors font-medium"
                >
                  Load NMIET Attendance Sheet
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
                  onChange={handleFileChange}
                  className="hidden"
                />
              </label>
            </div>

            {isProcessing && (
              <div className="text-center text-orange-400">
                <div className="inline-block animate-spin rounded-full h-6 w-6 border-b-2 border-orange-400 mr-2"></div>
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
                    Batch: {availableSheets.join(', ')}
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

          {/* Batch Selection - Similar to SRM Online */}
          {availableSheets.length > 0 && (
            <div className="mt-4 p-4 bg-orange-600/10 border border-orange-500/30 rounded-lg">
              <div className="flex items-center justify-between mb-3">
                <div>
                  <h3 className="text-sm font-semibold text-orange-300 mb-1">Batch Selection</h3>
                  <p className="text-xs text-orange-200/80">
                    NMIET MERN Stack Training - Single batch detected
                  </p>
                </div>
                <span className="bg-orange-600 text-white text-xs px-2 py-1 rounded-full">
                  1 batch
                </span>
              </div>
              
              {/* Single Sheet Display */}
              <div className="space-y-3">
                <div className="bg-gray-700/30 border border-gray-600/50 rounded-lg p-3">
                  <div className="flex items-center gap-3">
                    <div className="w-4 h-4 bg-orange-500 rounded-full"></div>
                    <div className="flex-1">
                      <div className="text-sm text-white font-medium">📊 MERN Stack</div>
                      <div className="text-xs text-orange-400">
                        NMIET Batch - Rows 5-95 (MERN Stack Training)
                      </div>
                    </div>
                    <div className="text-xs bg-green-600 text-white px-2 py-1 rounded">
                      Active
                    </div>
                  </div>
                </div>
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
                  <span className="text-xs px-2 py-1 bg-orange-600 text-white rounded">
                    NMIET
                  </span>
                  <span className="text-xs px-2 py-1 bg-green-600 text-white rounded">
                    1 batch active
                  </span>
                </div>
              </div>
              
              {/* Date Selection */}
              <div className="space-y-3">
                <label className="text-sm text-gray-300 font-medium">Select Date:</label>
                <select
                  value={selectedDate}
                  onChange={(e) => {
                    const newDate = e.target.value;
                    setSelectedDate(newDate);
                    const students = allSheetsData.get('Attendance') || [];
                    const stats = calculateAttendanceStats(newDate, students);
                    if (stats) {
                      setAttendanceStats(stats);
                      setAllSheetsAttendanceData(new Map([['Attendance', stats]]));
                    }
                  }}
                  className="w-full bg-gray-700 border border-gray-600 rounded-md px-3 py-2 text-white focus:ring-2 focus:ring-orange-500 focus:border-orange-500 select-auto"
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
                    <div className="bg-green-600/20 border border-green-500/30 rounded p-4 text-center">
                      <div className="text-2xl font-bold text-white">{attendanceStats.present}</div>
                      <div className="text-xs text-green-400">{attendanceStats.presentPercentage}%</div>
                      <div className="text-xs text-gray-400">Present</div>
                    </div>
                    <div className="bg-red-600/20 border border-red-500/30 rounded p-4 text-center">
                      <div className="text-2xl font-bold text-white">{attendanceStats.absent}</div>
                      <div className="text-xs text-red-400">{attendanceStats.absentPercentage}%</div>
                      <div className="text-xs text-gray-400">Absent</div>
                    </div>
                  </div>

                  {/* Progress Bar */}
                  <div className="space-y-2">
                    <div className="flex justify-between text-xs text-gray-400">
                      <span>Present: {attendanceStats.presentPercentage}%</span>
                      <span>Absent: {attendanceStats.absentPercentage}%</span>
                    </div>
                    <div className="w-full bg-gray-700 rounded-full h-2 overflow-hidden">
                      <div 
                        className="h-full bg-green-500 transition-all duration-500"
                        style={{ width: `${attendanceStats.presentPercentage}%` }}
                      ></div>
                    </div>
                  </div>
                </div>
              )}

              {/* Student Lists - Side by Side */}
              {attendanceStats && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  {/* Present Students */}
                  {attendanceStats.presentStudents.length > 0 ? (
                    <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-3 h-3 bg-green-500 rounded-full"></div>
                        <h4 className="text-sm font-semibold text-green-400">
                          Present Students ({attendanceStats.presentStudents.length})
                        </h4>
                      </div>
                      <div className="max-h-48 overflow-y-auto">
                        <div className="space-y-2">
                          {attendanceStats.presentStudents.map((student, index) => (
                            <div key={index} className="bg-gray-700/30 rounded p-3 border-l-2 border-green-500">
                              <div className="text-sm font-medium text-white">{student.name}</div>
                              <div className="text-xs text-gray-400">{student.email}</div>
                              {(student.course || student.branch) && (
                                <div className="text-xs text-green-300 mt-1">
                                  {student.course} {student.branch && `• ${student.branch}`}
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="bg-gray-700/20 border border-gray-600/30 rounded-lg p-4">
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-3 h-3 bg-gray-500 rounded-full"></div>
                        <h4 className="text-sm font-semibold text-gray-400">
                          Present Students (0)
                        </h4>
                      </div>
                      <div className="text-center py-4">
                        <p className="text-xs text-gray-500">No students were present</p>
                      </div>
                    </div>
                  )}

                  {/* Absent Students */}
                  {attendanceStats.absentStudents.length > 0 ? (
                    <div className="bg-red-600/10 border border-red-500/30 rounded-lg p-4">
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-3 h-3 bg-red-500 rounded-full"></div>
                        <h4 className="text-sm font-semibold text-red-400">
                          Absent Students ({attendanceStats.absentStudents.length})
                        </h4>
                      </div>
                      <div className="max-h-48 overflow-y-auto">
                        <div className="space-y-2">
                          {attendanceStats.absentStudents.map((student, index) => (
                            <div key={index} className="bg-gray-700/30 rounded p-3 border-l-2 border-red-500">
                              <div className="text-sm font-medium text-white">{student.name}</div>
                              <div className="text-xs text-gray-400">{student.email}</div>
                              {(student.course || student.branch) && (
                                <div className="text-xs text-red-300 mt-1">
                                  {student.course} {student.branch && `• ${student.branch}`}
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="bg-gray-700/20 border border-gray-600/30 rounded-lg p-4">
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-3 h-3 bg-gray-500 rounded-full"></div>
                        <h4 className="text-sm font-semibold text-gray-400">
                          Absent Students (0)
                        </h4>
                      </div>
                      <div className="text-center py-4">
                        <p className="text-xs text-gray-500">All students were present</p>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          ) : (
            <div className="text-center py-4">
              <p className="text-xs text-gray-400">Upload attendance sheet to see analysis</p>
            </div>
          )}
        </section>
      </div>

      {/* Row 5: Email Template Generator */}
      <section className="bg-gray-800/50 border border-gray-700/50 rounded-xl p-6">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-3">
            <Mail className="w-5 h-5 text-gray-400" />
            <h2 className="text-lg font-semibold text-white">Email Template Generator</h2>
            <span className="bg-orange-600 text-white text-xs px-2 py-1 rounded-full">
              Staff Template
            </span>
          </div>
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
                    onClick={async () => {
                      await navigator.clipboard.writeText(emailTemplate.to);
                      setCopiedEmailTo(true);
                      setTimeout(() => setCopiedEmailTo(false), 2000);
                    }}
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
                    onClick={async () => {
                      await navigator.clipboard.writeText(emailTemplate.cc);
                      setCopiedEmailCc(true);
                      setTimeout(() => setCopiedEmailCc(false), 2000);
                    }}
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
                      onClick={async () => {
                        await navigator.clipboard.writeText(emailTemplateSubject);
                        setCopiedEmailSubject(true);
                        setTimeout(() => setCopiedEmailSubject(false), 2000);
                      }}
                      className="flex items-center gap-1 px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors"
                    >
                      {copiedEmailSubject ? (
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

      {/* Row 5: Student Email Templates - Full Width like NIET */}
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
            Choose which batch topic to include in the email template table (will show only 1 row)
          </p>
          <select
            value={selectedSheet || ''}
            onChange={(e) => setSelectedSheet(e.target.value)}
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
              <div className="space-y-3">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">Subject:</span>
                  <input type="text" value={absentStudentEmailSubject} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentSubject} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentSubject ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">To:</span>
                  <input type="text" value={absentStudentEmailTo} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentTo} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentTo ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">CC:</span>
                  <input type="text" value={absentStudentEmailCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentCc} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentCc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">BCC:</span>
                  <input type="text" value={absentStudentEmailBCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyAbsentStudentBcc} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedAbsentStudentBcc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="mt-2 text-xs text-red-400 font-medium">
                  {attendanceStats ? `${attendanceStats.absent} absent students will be BCC'd` : 'No attendance data available'}
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

          <div className="bg-green-600/10 border border-green-500/30 rounded-lg p-4">
            <div className="flex items-center justify-between mb-4">
              <div>
                <div className='flex flex-row'>
                  <h3 className="text-lg font-semibold text-green-300 mb-2">Email for Present Students</h3>
                  {attendanceStats && (
                <div className="flex items-center gap-2 ml-4 -mt-2">
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
              <div className="space-y-3">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">Subject:</span>
                  <input type="text" value={presentStudentEmailSubject} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentSubject} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentSubject ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">To:</span>
                  <input type="text" value={presentStudentEmailTo} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentTo} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentTo ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">CC:</span>
                  <input type="text" value={presentStudentEmailCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentCc} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentCc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-400 w-12">BCC:</span>
                  <input type="text" value={presentStudentEmailBCC} readOnly className="flex-1 bg-gray-800 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300" />
                  <button onClick={copyPresentStudentBcc} className="flex items-center gap-1 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-medium rounded transition-colors">
                    {copiedPresentStudentBcc ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  </button>
                </div>
                <div className="mt-2 text-xs text-green-400 font-medium">
                  {attendanceStats ? `${attendanceStats.present} present students will be BCC'd` : 'No attendance data available'}
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
        
        <div className="mt-6 bg-orange-600/10 border border-orange-500/30 rounded-lg p-4">
          <div className="flex items-start gap-3">
            <Mail className="w-5 h-5 text-orange-400 mt-0.5 flex-shrink-0" />
            <div>
              <h4 className="text-orange-300 text-sm font-semibold mb-2">How to Use Student Email Templates</h4>
              <ul className="text-xs text-orange-200/80 space-y-1 list-disc list-inside">
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