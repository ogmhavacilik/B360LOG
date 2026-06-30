/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  Plane, 
  PlusCircle, 
  Search, 
  Users, 
  History, 
  Database,
  ArrowRight,
  Clock,
  MapPin,
  Flame,
  Camera,
  Satellite,
  Download,
  ExternalLink,
  AlertCircle,
  CheckCircle2,
  FileText,
  Trash2,
  Lock
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import { motion, AnimatePresence } from 'motion/react';
import XLSX from 'xlsx-js-style';
import { FlightLog, PILOTS, TECHNICIANS, GOREV_TIPLERI } from './types';

// Google Apps Script URLs
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwpqbW9sHqUZIUwGf4naJN_ZH0aCXvGEYSdqBN6mftyFVZGAofjrnMPfutcO5maBc4/exec';
const PILOT_DATA_URL = 'https://script.google.com/macros/s/AKfycbx5jJmZ6sU8qpSwCFwp41z_7fJYy-buB4BD5686jAoM3xqVw39m3q3iDbkVbQCDUZ5U/exec';
const TECH_DATA_URL = 'https://script.google.com/macros/s/AKfycbwcRIUfG0WfyEb5aWiJnpiNTvbo5XGz_WcetUlkoQGmZuTTBxvdsvsV2HQRPq8ewqEy/exec';
const PERSONNEL_SHEET_URL = 'https://script.google.com/macros/s/AKfycbytVmnCY7Spjg-Rges0k-BgEJqZSM8iNJoXu0UHuKFeubm4vlSzemzzVec4UgUQ96Q7/exec';

interface PersonData {
  fullName: string;
  photoUrl: string;
  role?: string;
  title?: string;
}

// Utility to calculate flight duration
const calculateDuration = (kalkis: string, inis: string): string => {
  if (!kalkis || !inis) return '00:00';
  const [h1, m1] = kalkis.split(':').map(Number);
  const [h2, m2] = inis.split(':').map(Number);
  
  let diffMin = (h2 * 60 + m2) - (h1 * 60 + m1);
  if (diffMin < 0) diffMin += 1440; // Wrap around midnight
  
  const h = Math.floor(diffMin / 60);
  const m = diffMin % 60;
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
};

// Utility to handle Excel ISO date/time format (e.g. 1899-12-30T01:50:00.000Z)
const normalize = (str: string): string => {
  if (!str) return '';
  try {
    let s = str.toString()
      .replace(/İ/g, "I")
      .replace(/ı/g, "i")
      .replace(/Ğ/g, "G")
      .replace(/ğ/g, "g")
      .replace(/Ü/g, "U")
      .replace(/ü/g, "u")
      .replace(/Ş/g, "S")
      .replace(/ş/g, "s")
      .replace(/Ö/g, "O")
      .replace(/ö/g, "o")
      .replace(/Ç/g, "C")
      .replace(/ç/g, "c")
      .toUpperCase();

    // Word boundary clean for roles to avoid partial name match
    const roles = ["OPERATOR", "OPERATÖR", "TEKNISYEN", "TEKNİSYEN", "KAPTAN", "PILOT", "PİLOT", "OPR", "TEK", "EKIP", "EKİP", "PERSONEL"];
    roles.forEach(role => {
      const re = new RegExp('\\b' + role + '\\b', 'g');
      s = s.replace(re, ' ');
    });

    return s
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  } catch (e) {
    return str.toString().toUpperCase().trim();
  }
};

const getSurname = (fullName: string): string => {
  if (!fullName) return '';
  const parts = fullName.trim().split(/\s+/);
  return parts[parts.length - 1].toUpperCase();
};

const getDriveThumbnail = (url: string) => {
  if (!url || !url.includes('google.com')) {
    if (url?.startsWith('http')) return url;
    return '';
  }
  
  const idMatch = url.match(/(?:id=|\/d\/|\/file\/d\/|\/u\/\d+\/d\/|\/uc\?id=)([a-zA-Z0-9-_]{25,})/);
  if (idMatch && idMatch[1]) {
    return `https://drive.google.com/thumbnail?sz=400&id=${idMatch[1]}`;
  }
  return url;
};

const isPersonMatch = (logCellValue: any, selectedFullName: string): boolean => {
  if (logCellValue === undefined || logCellValue === null || !selectedFullName) return false;
  
  const cellStr = logCellValue.toString().trim();
  if (!cellStr) return false;

  // Normalizes Turkish characters and casings to standardize comparison
  const normalizeForMatch = (str: string): string => {
    return str
      .replace(/İ/g, "I")
      .replace(/ı/g, "I")
      .replace(/Ğ/g, "G")
      .replace(/ğ/g, "G")
      .replace(/Ü/g, "U")
      .replace(/ü/g, "U")
      .replace(/Ş/g, "S")
      .replace(/ş/g, "S")
      .replace(/Ö/g, "O")
      .replace(/ö/g, "O")
      .replace(/Ç/g, "C")
      .replace(/ç/g, "C")
      .toUpperCase()
      .replace(/KEBABCI/g, "KEBAPCI")
      .trim();
  };

  const cleanCell = normalizeForMatch(cellStr);
  const cleanTarget = normalizeForMatch(selectedFullName);

  if (cleanCell === cleanTarget) return true;

  // Split clean text blocks into arrays of words, removing all punctuation
  const getWords = (s: string) => {
    return s
      .replace(/[^\w\s]/g, ' ')
      .split(/\s+/)
      .filter(Boolean);
  };

  const cellWords = getWords(cleanCell);
  const targetWords = getWords(cleanTarget);

  if (cellWords.length === 0 || targetWords.length === 0) return false;

  // Take the last word as the surname of the selected search target
  const targetSurname = targetWords[targetWords.length - 1];

  // Returns true only if the surname exists as a full standalone word in the cell words
  if (cellWords.includes(targetSurname)) {
    return true;
  }

  return false;
};

const formatDateToTR = (dateStr: string): string => {
  if (!dateStr || dateStr === 'null') return '';
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return `${parts[2]}.${parts[1]}.${parts[0]}`;
  }
  return dateStr;
};

// Utility to parse numeric values handling different locale formats
const parseNumeric = (val: any): number => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  
  let str = val.toString().trim();
  if (!/[0-9]/.test(str)) return 0;

  // Handle string additions like "250+367" or "343+350"
  if (str.includes('+')) {
    const parts = str.split('+');
    return parts.reduce((acc, part) => acc + parseNumeric(part), 0);
  }

  // Clear spaces that might be used as a thousands separator
  str = str.replace(/\s/g, '');

  // Remove any unit suffix letters (e.g. "DK", "HK", "ADET", "Col") to avoid interference
  str = str.replace(/[^0-9.,-]/g, '');

  // Case 1: Has both dot and comma (e.g. "1.289.850,50" or "1,289,850.50")
  if (str.includes(',') && str.includes('.')) {
    const lastComma = str.lastIndexOf(',');
    const lastDot = str.lastIndexOf('.');
    if (lastComma > lastDot) {
      // Turkish format: "1.289.850,50" -> remove dots, replace comma with dot
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      // English format: "1,289,850.50" -> remove commas
      str = str.replace(/,/g, '');
    }
  } 
  // Case 2: Only has a comma (e.g. "855350,50" or "1,289,850")
  else if (str.includes(',')) {
    const commaCount = (str.match(/,/g) || []).length;
    if (commaCount > 1) {
      // Multiple commas -> thousands separator -> remove them
      str = str.replace(/,/g, '');
    } else {
      // Single comma -> decimal separator in Turkish -> replace with dot
      str = str.replace(',', '.');
    }
  } 
  // Case 3: Only has a dot (e.g. "1.289.850" or "855.350" or "123.45")
  else if (str.includes('.')) {
    const dotCount = (str.match(/\./g) || []).length;
    if (dotCount > 1) {
      // Multiple dots -> thousands separator -> remove them
      str = str.replace(/\./g, '');
    } else {
      // Single dot -> could be a thousands separator (Turkish binlik) or decimal (English)
      // If there are exactly 3 digits after the dot and the number can be large, it's a thousands separator
      const parts = str.split('.');
      if (parts[1] && parts[1].length === 3) {
        str = str.replace(/\./g, '');
      }
    }
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
};

// Helper to sum Uydu (Dk) while complying with Excel's treatment of text entries
const parseUyduDkForSum = (val: any): number => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const str = val.toString().trim();
  const parsed = parseNumeric(str);
  return isNaN(parsed) ? 0 : parsed;
};

const formatTimeValue = (val: any): string => {
  if (!val) return '00:00';
  const str = val.toString().trim();
  if (str.includes('T')) {
    try {
      const timePart = str.split('T')[1] || '';
      const parts = timePart.split(':');
      if (parts.length < 2) return '00:00';
      return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
    } catch (e) {
      return str;
    }
  }
  return str;
};

// Utility to format date for display (GG.AY.YIL)
const formatDisplayDate = (dateStr: string): string => {
  if (!dateStr) return '-';
  try {
    const [year, month, day] = dateStr.split('-');
    if (!year || !month || !day) return dateStr;
    return `${day}.${month}.${year}`;
  } catch (e) {
    return dateStr;
  }
};

// Utility to calculate total flight hours for a person across ALL columns (Surname search compatible)
const calculateTotalHours = (logs: FlightLog[], query: string): string => {
  let totalMinutes = 0;
  if (!query) return '00:00';

  logs.forEach(log => {
    // Check every possible personal slot
    const crew = [
      log.kaptanPilot, log.ikinciPilot, 
      log.teknisyen1, log.teknisyen2, 
      log.operator1, log.operator2
    ];

    if (crew.some(name => isPersonMatch(name || '', query))) {
      const duration = log.ucusSuresi || '00:00';
      const [h, m] = duration.split(':').map(Number);
      if (!isNaN(h) && !isNaN(m)) {
        totalMinutes += (h * 60 + m);
      }
    }
  });

  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
};

// Utility to get person stats (total count, role, distribution, photo)
const getPersonStats = (personName: string, logs: FlightLog[], personnelData: Record<string, PersonData>, defaultType: string) => {
  let pilotOccurrences = 0;
  let technicianOccurrences = 0;
  let operatorOccurrences = 0;
  
  let totalYanginHektar = 0;
  let totalGorevHektar = 0;
  let totalUyduDk = 0;
  let totalCekim = 0;

  const missionDistribution: Record<string, number> = {};
  
  const matchedLogs = logs.filter(log => {
    const isPilot = isPersonMatch(log.kaptanPilot || '', personName) || isPersonMatch(log.ikinciPilot || '', personName);
    const isTech = isPersonMatch(log.teknisyen1 || '', personName) || isPersonMatch(log.teknisyen2 || '', personName);
    const isOp = isPersonMatch(log.operator1 || '', personName) || isPersonMatch(log.operator2 || '', personName);
    
    if (isPilot) pilotOccurrences++;
    if (isTech) technicianOccurrences++;
    if (isOp) operatorOccurrences++;

    const isAnyMatch = isPilot || isTech || isOp;
    if (isAnyMatch) {
      const type = log.gorevTipi || 'Diğer';
      missionDistribution[type] = (missionDistribution[type] || 0) + 1;
      
      totalYanginHektar += Number(log.k9YanginHektar || 0);
      totalGorevHektar += Number(log.tk9GorevHektar || 0);
      totalUyduDk += parseUyduDkForSum(log.uyduDk);
      totalCekim += Number(log.miktarCekim || 0);
    }
    return isAnyMatch;
  });

  let role = defaultType === 'Pilot' ? 'Pilot' : 'Uçuş Teknisyeni';
  if (operatorOccurrences > (technicianOccurrences + pilotOccurrences)) {
    role = 'Operatör';
  } else if (technicianOccurrences > (pilotOccurrences + operatorOccurrences)) {
    role = 'Uçuş Teknisyeni';
  } else if (pilotOccurrences > (technicianOccurrences + operatorOccurrences)) {
    role = 'Pilot';
  }

  const pData = personnelData[normalize(personName)];
  let finalPhoto = pData?.photoUrl;
  let finalTitle = pData?.title;
  let finalRole = pData?.role;

  if (!finalPhoto) {
    const partialMatch = Object.entries(personnelData).find(([key, val]) => 
      normalize(personName).includes(key) || key.includes(normalize(personName))
    );
    if (partialMatch) {
      finalPhoto = partialMatch[1].photoUrl;
      finalTitle = partialMatch[1].title;
      finalRole = partialMatch[1].role;
    }
  }

  return {
    count: matchedLogs.length,
    role: finalRole || role,
    distribution: Object.entries(missionDistribution).sort((a, b) => b[1] - a[1]),
    photoUrl: finalPhoto,
    title: finalTitle,
    stats: {
      totalHektar: totalYanginHektar + totalGorevHektar,
      yanginHektar: totalYanginHektar,
      gorevHektar: totalGorevHektar,
      uyduDk: totalUyduDk,
      cekim: totalCekim
    }
  };
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'home' | 'add' | 'search' | 'personnel'>('home');
  const [logs, setLogs] = useState<FlightLog[]>([]);
  const [personnelData, setPersonnelData] = useState<Record<string, PersonData>>(() => {
    const initial: Record<string, PersonData> = {};
    PILOTS.forEach(name => {
      initial[normalize(name)] = {
        fullName: name,
        photoUrl: '',
        role: 'Pilot',
        title: 'B-360 PİLOT'
      };
    });
    TECHNICIANS.forEach(name => {
      initial[normalize(name)] = {
        fullName: name,
        photoUrl: '',
        role: 'Teknisyen',
        title: 'B-360 TEKNİSYEN'
      };
    });
    return initial;
  });
  const [isLoading, setIsLoading] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isSuccess, setIsSuccess] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState('');
  const [isPasswordVerified, setIsPasswordVerified] = useState(false);

  const handleOpenKayitlar = () => {
    if (isPasswordVerified) {
      setActiveTab('search');
    } else {
      setPasswordInput('');
      setPasswordError('');
      setShowPasswordModal(true);
    }
  };

  const handleVerifyPassword = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (passwordInput === '360') {
      setIsPasswordVerified(true);
      setShowPasswordModal(false);
      setActiveTab('search');
    } else {
      setPasswordError('Hatalı Şifre! Lütfen tekrar deneyin.');
    }
  };
  const [searchTerm, setSearchTerm] = useState('');
  const [personnelSearch, setPersonnelSearch] = useState('');
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [filters, setFilters] = useState(() => {
    const currentYear = new Date().getFullYear();
    return {
      pilot: '',
      bolge: '',
      tip: '',
      startDate: `${currentYear}-01-01`,
      endDate: `${currentYear + 1}-01-01`
    };
  });

  const [personnelFilters, setPersonnelFilters] = useState(() => {
    const currentYear = new Date().getFullYear();
    return {
      startDate: `${currentYear}-01-01`,
      endDate: `${currentYear + 1}-01-01`
    };
  });

  const pilotsInResults = useMemo(() => {
    const list = new Set([...PILOTS]);
    Object.values(personnelData).forEach((p: PersonData) => {
      const isPilot = p.role?.toUpperCase().includes('PİLOT') || 
                      p.role?.toUpperCase().includes('PILOT') || 
                      p.role?.toUpperCase().includes('KAPTAN') || 
                      p.title?.toUpperCase().includes('PİLOT') || 
                      p.title?.toUpperCase().includes('PILOT');
      if (isPilot) {
        list.add(p.fullName);
      }
    });
    return Array.from(list).sort();
  }, [personnelData]);

  const technicalInResults = useMemo(() => {
    const list = new Set([...TECHNICIANS]);
    Object.values(personnelData).forEach((p: PersonData) => {
      const isPilot = p.role?.toUpperCase().includes('PİLOT') || 
                      p.role?.toUpperCase().includes('PILOT') || 
                      p.role?.toUpperCase().includes('KAPTAN') || 
                      p.title?.toUpperCase().includes('PİLOT') || 
                      p.title?.toUpperCase().includes('PILOT');
      if (!isPilot) {
        list.add(p.fullName);
      }
    });
    return Array.from(list).sort();
  }, [personnelData]);

  const filteredLogsForPersonnel = useMemo(() => {
    return logs.filter(log => {
      const matchesStartDate = !personnelFilters.startDate || (log.tarih && log.tarih >= personnelFilters.startDate);
      const matchesEndDate = !personnelFilters.endDate || (log.tarih && log.tarih <= personnelFilters.endDate);
      return matchesStartDate && matchesEndDate;
    });
  }, [logs, personnelFilters]);

  // Fetch extra personnel info (photos)
  const fetchPersonnelInfo = async () => {
    try {
      const getData = async (url: string) => {
        try {
          const r = await fetch(url);
          if (!r.ok) return null;
          const txt = await r.text();
          try {
            return JSON.parse(txt);
          } catch (e) {
            const jsonMatch = txt.match(/\[.*\]|\{.*\}/s);
            return jsonMatch ? JSON.parse(jsonMatch[0]) : null;
          }
        } catch (e) {
          return null;
        }
      };

      const [pilotRes, techRes, personnelRes] = await Promise.all([
        getData(PILOT_DATA_URL),
        getData(TECH_DATA_URL),
        getData(PERSONNEL_SHEET_URL)
      ]);

      setPersonnelData(prev => {
        const dataMap: Record<string, PersonData> = { ...prev };
        
        const processItem = (p: any) => {
          if (!p) return;
          
          let name = '';
          let photo = '';
          let role = '';
          let title = '';

          if (Array.isArray(p)) {
            if (p.length < 2) return;
            const firstVal = p[0]?.toString().toUpperCase();
            if (firstVal === 'ID' || firstVal === 'AD SOYAD' || p.some(v => v?.toString().toUpperCase() === 'FOTOĞRAF')) return;

            // Search for photo URL
            photo = p.find(val => typeof val === 'string' && (val.includes('drive.google.com') || val.includes('http'))) || '';
            
            // Search for Name
            name = p.find((val) => {
              if (typeof val !== 'string' || val.length < 3 || val.includes('http')) return false;
              const up = val.toUpperCase();
              if (['PİLOT', 'TEKNİSYEN', 'KAPTAN', 'OPERATÖR', 'ŞOFÖR', 'MEMUR', 'İŞÇİ'].includes(up)) return false;
              return true;
            }) || '';
            
            role = p.find(val => {
              if (typeof val !== 'string') return false;
              const up = val.toUpperCase();
              return ['PİLOT', 'TEKNİSYEN', 'OPERATÖR', 'ŞÖFÖR', 'İŞÇİ', 'MEMUR', 'DİĞER'].some(r => up.includes(r));
            }) || '';

            title = p.find(val => {
               if (typeof val !== 'string' || val.length < 3 || val.includes('http')) return false;
               const up = val.toUpperCase();
               if (['KAPTAN', 'SİSTEM', 'TEKNİK', 'TEKNİSYEN', 'TEKNİKER', 'MÜHENDİS', 'ŞÖFÖR', 'İŞÇİ', 'MEMUR'].some(kw => up.includes(kw))) return true;
               return false;
            }) || '';
          } else {
            name = p.AD_SOYAD || p["AD SOYAD"] || p.fullName || p.FULL_NAME || p.AdSoyad || p.adSoyad || p.Name || p.name || '';
            photo = p.PHOTO_URL || p["FOTOĞRAF"] || p.photoUrl || p.PHOTO || p.photo || p.fotograf || '';
            role = p.ROLE || p["GÖREV"] || p.role || p.gorev || p.GÖREV || '';
            title = p.TITLE || p["ÜNVAN"] || p.title || p.unvan || p.ÜNVAN || '';
          }

          if (name && typeof name === 'string' && name.length > 2) {
            const cleanName = name.replace(/\s*\(.*?\)\s*/g, '').trim();
            const normKey = normalize(cleanName);
            const allowedNames = new Set([...PILOTS, ...TECHNICIANS].map(n => normalize(n)));
            if (allowedNames.has(normKey)) {
              dataMap[normKey] = {
                fullName: cleanName,
                photoUrl: photo ? getDriveThumbnail(typeof photo === 'string' ? photo : '') : (dataMap[normKey]?.photoUrl || ''),
                role: (typeof role === 'string' ? role : '') || (dataMap[normKey]?.role || ''),
                title: (typeof title === 'string' ? title : '') || (dataMap[normKey]?.title || '')
              };
            }
          }
        };

        const processData = (input: any) => {
          if (!input) return;
          const arr = Array.isArray(input) ? input : (input.data && Array.isArray(input.data) ? input.data : []);
          arr.forEach(item => processItem(item));
        };

        processData(pilotRes);
        processData(techRes);
        processData(personnelRes);

        return dataMap;
      });
    } catch (err) {
      console.warn('Personnel photo data could not be fetched:', err);
    }
  };

  const fetchExternalData = async () => {
    setIsLoading(true);
    setSyncError(null);
    try {
      // Step 1: Try to fetch the master sheet as raw Excel Binary (.xlsx)
      // This is 100% immune to duplicate column headers (like Teknisyen / Operator)
      // because it parses the raw grid values by exact cell index, not object keys.
      try {
        const response = await fetch(`${GAS_URL}?action=download&gid=1476570479`);
        if (response.ok) {
          const result = await response.json();
          if (result && result.status === 'success' && result.base64) {
            const byteCharacters = atob(result.base64);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {
              byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            const byteArray = new Uint8Array(byteNumbers);
            
            // Read workbook via SheetJS
            const workbook = XLSX.read(byteArray, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert worksheet to an Array of Arrays (AOA) with headers intact
            const rawAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (rawAOA && Array.isArray(rawAOA) && rawAOA.length > 0) {
              processImportedData(rawAOA);
              setIsLoading(false);
              return; // Successfully processed raw grid with zero key collision!
            }
          }
        }
      } catch (excelError) {
        console.warn('Failed to fetch Excel binary, falling back to JSON get:', excelError);
      }

      // Step 2: Fallback to standard doGet JSON response if Excel gets blocked or fails
      const url = GAS_URL;
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
      
      const data = await response.json();
      processImportedData(data);
    } catch (e) {
      console.error('Veri çekme hatası:', e);
      setSyncError('BAĞLANTI HATASI: Google Apps Script "Herkes" erişimine açık olmayabilir veya tarayıcı erişimi engelliyor. Lütfen sayfayı yenileyip tekrar deneyin.');
    } finally {
      setIsLoading(false);
    }
  };

  const processImportedData = (data: any, append: boolean = false) => {
    if (Array.isArray(data)) {
      setLogs(prev => {
        const processedList: FlightLog[] = [];
        const isAOA = data.length > 0 && Array.isArray(data[0]);

        // Parse Excel Date Serial (e.g. 44100) or standard ISO/Turkish Date format
        const parseExcelDate = (val: any): string => {
          if (val === undefined || val === null || val === '') return '';
          const str = val.toString().trim();
          
          if (/^\d{5}$/.test(str)) {
            const dateNum = Number(str);
            const excelDate = new Date((dateNum - 25569) * 86400 * 1000);
            return excelDate.toISOString().split('T')[0];
          }
          
          if (str.includes('T')) {
            return str.split('T')[0];
          }
          
          const parts = str.split(/[:./-]/);
          if (parts.length === 3) {
            let p0 = parts[0].trim().padStart(2, '0');
            let p1 = parts[1].trim().padStart(2, '0');
            let p2 = parts[2].trim();
            if (p0.length === 4) {
              return `${p0}-${p1}-${p2}`;
            } else {
              if (p2.length === 2) {
                p2 = p2.startsWith('2') ? `20${p2}` : `19${p2}`;
              }
              return `${p2}-${p1}-${p0}`;
            }
          }
          return str;
        };

        // Parse Excel Time fraction (e.g. 0.58333) or HH:mm format
        const parseExcelTime = (val: any): string => {
          if (val === undefined || val === null || val === '') return '00:00';
          const str = val.toString().trim();
          
          const num = Number(str);
          if (!isNaN(num) && num > 0 && num < 1) {
            const totalMin = Math.round(num * 1440);
            const h = Math.floor(totalMin / 60);
            const m = totalMin % 60;
            return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
          }

          if (str.includes('T')) {
            try {
              const timePart = str.split('T')[1] || '';
              const parts = timePart.split(':');
              if (parts.length >= 2) {
                return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
              }
            } catch (e) {
              // Ignore
            }
          }

          const parts = str.split(':');
          if (parts.length >= 2) {
            const h = parts[0].trim().padStart(2, '0');
            const m = parts[1].trim().padStart(2, '0');
            return `${h}:${m}`;
          }

          return str;
        };

        data.forEach((l, rowIdx) => {
          if (!l) return;
          
          // Skip header row if AOA format
          if (isAOA && rowIdx === 0) {
            const firstCell = (l[0] || '').toString().toLowerCase();
            if (firstCell.includes('sıra') || firstCell.includes('tarih') || firstCell.includes('no')) {
              return;
            }
          }

          const val = (keys: string[], colIdx?: number) => {
            if (isAOA && Array.isArray(l) && colIdx !== undefined) {
              if (l[colIdx] !== undefined && l[colIdx] !== null && l[colIdx] !== '') return l[colIdx];
            }
            if (l && typeof l === 'object' && !Array.isArray(l)) {
              // Priority 1: Check column index key (e.g. col_4, col_5) which is unique and collision-immune!
              if (colIdx !== undefined) {
                const indexKey = 'col_' + colIdx;
                if (l[indexKey] !== undefined && l[indexKey] !== null && l[indexKey] !== '') return l[indexKey];
              }
              // Priority 2: Check standard strings
              for (const k of keys) {
                if (l[k] !== undefined && l[k] !== null && l[k] !== '') return l[k];
                const target = k.toLowerCase().replace(/[^\w]/g, '');
                for (const actualKey in l) {
                  const actualClean = actualKey.toLowerCase().replace(/[^\w]/g, '');
                  if (actualClean === target) {
                     if (l[actualKey] !== undefined && l[actualKey] !== null && l[actualKey] !== '') return l[actualKey];
                  }
                }
              }
            }
            return '';
          };

          const idRaw = (val(['Sıra No', 'id'], 0) || '').toString().trim();
          if (!idRaw) return;

          const rawTarih = val(['Tarih', 'tarih'], 1);
          const historyTarih = parseExcelDate(rawTarih);

          // 1:1 direct column mapping from Excel to FlightLog attributes without changing any text
          const cleaned: FlightLog = {
            id: idRaw,
            tarih: historyTarih,
            kaptanPilot: (val(['Kaptan Pilot', 'kaptanPilot'], 2) || '').toString().trim(),
            ikinciPilot: (val(['2. Pilot', 'ikinciPilot'], 3) || '').toString().trim(),
            teknisyen1: (val(['Teknisyen 1', 'teknisyen1', 'teknisyen', 'Teknisyen'], 4) || '').toString().trim(),
            operator1: (val(['Operator 1', 'operator1', 'operator', 'Operator', 'operatör', 'Operatör'], 5) || '').toString().trim(),
            teknisyen2: (val(['Teknisyen 2', 'teknisyen2'], 6) || '').toString().trim(),
            operator2: (val(['Operator 2', 'operator2', 'Operatör 2'], 7) || '').toString().trim(),
            gorevTipi: (val(['Görev Tipi', 'gorevTipi'], 8) || '').toString().trim(),
            gorevBolgesi: (val(['Görev Bölgesi', 'gorevBolgesi'], 9) || '').toString().trim(),
            kalkis: parseExcelTime(val(['Kalkış', 'kalkis'], 10)),
            inis: parseExcelTime(val(['İniş', 'inis'], 11)),
            ucusSuresi: parseExcelTime(val(['Uçuş Süresi', 'Süre', 'ucusSuresi'], 12)) || '00:00',
            k9YanginHektar: parseNumeric(val(['TK-9 Yangın(Hektar)', 'YANGIN(Hk)', 'k9YanginHektar'], 13)),
            miktarCekim: parseNumeric(val(['Miktar (Çekim)', 'ÇEKİM', 'miktarCekim'], 14)),
            tk9GorevHektar: parseNumeric(val(['TK-9 gorev(Hektar)', 'GÖREV(Hk)', 'tk9GorevHektar'], 15)),
            uyduDk: (() => {
              const rawVal = val(['Uydu (Dk)', 'UYDU(Dk)', 'uyduDk'], 16);
              if (rawVal === undefined || rawVal === null) return 0;
              if (typeof rawVal === 'string') {
                const trimmed = rawVal.trim();
                if (trimmed.includes('+')) return trimmed;
                const parsed = parseNumeric(trimmed);
                return isNaN(parsed) ? trimmed : parsed;
              }
              return rawVal;
            })(),
            aciklama: (val(['AÇIKLAMA', 'aciklama'], 17) || '').toString().trim()
          };
          
          processedList.push(cleaned);
        });

        const list = append ? [...prev, ...processedList] : processedList;
        
        // Remove duplicates on id AND tarih to maintain data consistency
        const uniqueMap = new Map<string, FlightLog>();
        list.forEach((item) => {
          const key = `${item.tarih}_${item.id}`;
          uniqueMap.set(key, item);
        });

        return Array.from(uniqueMap.values()).sort((a, b) => {
          const idA = parseInt(a.id) || 0;
          const idB = parseInt(b.id) || 0;
          if (idA !== idB) return idB - idA; // Newest first
          return b.tarih.localeCompare(a.tarih);
        });
      });
      setSyncError(null);
    } else {
      setSyncError('Geçersiz veri formatı. Veri bir liste (array) olmalıdır.');
    }
  };

  // Load from localStorage on mount and try external fetch
  useEffect(() => {
    const saved = localStorage.getItem('b360_flight_logs');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          // Deduplicate existing state on load just in case
          const logMap = new Map<string, FlightLog>();
          parsed.forEach(l => {
            if (l && l.id && l.tarih) {
              const k = `${l.tarih}_${l.id}`;
              logMap.set(k, l);
            }
          });
          setLogs(Array.from(logMap.values()).sort((a, b) => {
            const idA = parseInt(a.id) || 0;
            const idB = parseInt(b.id) || 0;
            if (idA !== idB) return idA - idB;
            return a.tarih.localeCompare(b.tarih);
          }));
        } else {
          setLogs([]);
        }
      } catch (e) {
        console.error('Failed to load logs', e);
        setLogs([]);
      }
    } else {
      setLogs([]);
    }
    // Auto-fetch external data on start
    fetchExternalData();
    fetchPersonnelInfo();
  }, []);

  // Save to localStorage whenever logs change
  useEffect(() => {
    localStorage.setItem('b360_flight_logs', JSON.stringify(logs));
  }, [logs]);

  const addLog = async (log: Omit<FlightLog, 'id' | 'ucusSuresi'>) => {
    setIsSubmitting(true);
    // Sıra No mantığı: Excel'den gelenler bittikten sonra 281'den devam etsin
    const numericIds = logs.map(l => parseInt(l.id)).filter(n => !isNaN(n));
    const maxExistingId = numericIds.length > 0 ? Math.max(...numericIds) : 280;
    
    // Sistem üzerinden eklenecek ilk veri 281 olması için
    const nextIdNum = Math.max(maxExistingId, 280) + 1;
    const nextId = nextIdNum.toString();

    const ucusSuresi = calculateDuration(log.kalkis, log.inis);

    // Standardize names: We keep FULL NAMES for Google Sheets sync
    // but we can use getSurname for local display in some parts if needed.
    // However, the user said "appropriate format for excel pages", which usually means full names.
    const newLog: FlightLog = {
      ...log,
      k9YanginHektar: log.gorevTipi === 'Yangın Uçuşu' ? log.k9YanginHektar : 0,
      id: nextId,
      ucusSuresi: ucusSuresi
    };

    setLogs([newLog, ...logs]);

    // Sync with Sheets
    try {
      // Use no-cors to avoid preflight issues with Google Apps Script
      // JSON data is sent as text/plain to stay as a "simple request"
      await fetch(GAS_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'text/plain',
        },
        body: JSON.stringify(newLog),
      });
      
      // Since no-cors doesn't allow reading the response, 
      // and network errors usually throw, we provide positive feedback
      setIsSuccess(true);
      setTimeout(() => {
        setIsSuccess(false);
        setActiveTab('home');
      }, 3000);
    } catch (e) {
      console.error('Auto-sync failed', e);
      setSyncError('Veri e-tabloya gönderilemedi. Lütfen bağlantıyı kontrol edin.');
      alert('HATA: Veri Google E-Tabloya gönderilemedi. Lütfen script ayarlarını veya internetinizi kontrol edin.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const [editingLog, setEditingLog] = useState<FlightLog | null>(null);

  const deleteLog = async (id: string) => {
    if (!window.confirm(`Sıra No ${id} olan kaydı silmek istediğinizden emin misiniz?`)) {
      return;
    }
    
    setIsLoading(true);
    const previousLogs = [...logs];
    setLogs(prev => prev.filter(l => l.id !== id));
    
    try {
      await fetch(GAS_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'text/plain',
        },
        body: JSON.stringify({
          action: 'delete',
          ids: [id]
        }),
      });
      alert('Kayıt başarıyla silindi (E-tablo ile senkronize edildi).');
    } catch (e) {
      console.error('Delete sync failed', e);
      setLogs(previousLogs);
      alert('HATA: Kayıt e-tablodan silinemedi. Lütfen internetinizi ve script ayarlarını kontrol edin.');
    } finally {
      setIsLoading(false);
    }
  };

  const updateLog = async (id: string, updatedFields: Omit<FlightLog, 'id' | 'ucusSuresi'>) => {
    setIsSubmitting(true);
    
    const ucusSuresi = calculateDuration(updatedFields.kalkis, updatedFields.inis);
    const updatedLog: FlightLog = {
      ...updatedFields,
      k9YanginHektar: updatedFields.gorevTipi === 'Yangın Uçuşu' ? updatedFields.k9YanginHektar : 0,
      id: id,
      ucusSuresi: ucusSuresi
    };

    const previousLogs = [...logs];
    setLogs(prev => prev.map(l => l.id === id ? updatedLog : l));

    try {
      await fetch(GAS_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'text/plain',
        },
        body: JSON.stringify({
          action: 'update',
          id: id,
          data: updatedLog
        }),
      });
      alert('Kayıt başarıyla güncellendi (E-tablo ile senkronize edildi).');
      setEditingLog(null);
    } catch (e) {
      console.error('Update sync failed', e);
      setLogs(previousLogs);
      alert('HATA: Kayıt güncellenemedi. Lütfen internetinizi ve script ayarlarını kontrol edin.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const filteredLogs = useMemo(() => {
    return logs.filter(log => {
      const s = searchTerm.toLowerCase().trim();
      
      // Better text search across all visible data
      const searchData = [
        log.id,
        log.tarih,
        log.kaptanPilot,
        log.ikinciPilot,
        log.teknisyen1,
        log.operator1,
        log.teknisyen2,
        log.operator2,
        log.gorevTipi,
        log.gorevBolgesi,
        log.aciklama
      ].map(v => (v || '').toString().toLowerCase());

      const matchesSearch = !s || searchData.some(val => val.includes(s));
      
      const matchesPilot = !filters.pilot || [
        log.id,
        log.tarih,
        log.kaptanPilot,
        log.ikinciPilot, 
        log.teknisyen1,
        log.operator1, 
        log.teknisyen2,
        log.operator2,
        log.gorevTipi,
        log.gorevBolgesi,
        log.kalkis,
        log.inis,
        log.ucusSuresi,
        log.k9YanginHektar,
        log.miktarCekim,
        log.tk9GorevHektar,
        log.uyduDk,
        log.aciklama
      ].some(val => isPersonMatch(val, filters.pilot));

      const matchesBolge = !filters.bolge || (log.gorevBolgesi && log.gorevBolgesi.toLowerCase().includes(filters.bolge.toLowerCase()));
      const matchesTip = !filters.tip || log.gorevTipi === filters.tip;
      
      const matchesStartDate = !filters.startDate || (log.tarih && log.tarih >= filters.startDate);
      const matchesEndDate = !filters.endDate || (log.tarih && log.tarih <= filters.endDate);

      return matchesSearch && matchesPilot && matchesBolge && matchesTip && matchesStartDate && matchesEndDate;
    }).sort((a, b) => {
      const idA = parseInt(a.id) || 0;
      const idB = parseInt(b.id) || 0;
      if (idA !== idB) return idB - idA; // Consistent with import sort
      return b.tarih.localeCompare(a.tarih);
    });
  }, [logs, searchTerm, filters]);

  const exportToExcel = async (isRawExport = false) => {
    let listToExport = filteredLogs;
    let fileName = `B360_Ucus_Kayitlari_${new Date().toISOString().split('T')[0]}.xlsx`;

    if (isRawExport) {
      setIsLoading(true);
      try {
        // GAS üzerinden doğrudan Excel Binary verisini alıyoruz (Orijinal Tasarım İçin)
        const response = await fetch(`${GAS_URL}?action=download&gid=1476570479`);
        if (!response.ok) throw new Error('Sunucuya bağlanılamadı');
        
        const result = await response.json();
        if (result.status === 'success' && result.base64) {
          const byteCharacters = atob(result.base64);
          const byteNumbers = new Array(byteCharacters.length);
          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
          }
          const byteArray = new Uint8Array(byteNumbers);
          const blob = new Blob([byteArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
          
          const link = document.createElement('a');
          const url = URL.createObjectURL(blob);
          link.href = url;
          link.download = `OGM_B360_HAM_VERI_${new Date().toISOString().split('T')[0]}.xlsx`;
          document.body.appendChild(link);
          link.click();
          setTimeout(() => {
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
          }, 100);
          return; // Do not continue with local generation
        } else {
          throw new Error(result.message || 'Veri alınamadı');
        }
      } catch (e) {
        console.error('Ham veri indirme hatası:', e);
        alert('HATA: Ham veri sunucudan indirilemedi. Lütfen Apps Script yetkilerini onayladığınızdan emin olun.');
        return;
      } finally {
        setIsLoading(false);
      }
    }

    // Filtered/Raw Export with AOA (Array of Arrays) to match image exactly
    const headers = [
      "Sıra No (A)", "Tarih (B)", "Kaptan Pilot (C)", "2. Pilot (D)", "Teknisyen (E)", "Operator (F)", 
      "Teknisyen (G)", "Operator (H)", "Görev Tipi (I)", "Görev Bölgesi (J)", "Kalkış (K)", "İniş (L)", 
      "Uçuş Süresi (M)", "TK-9 Yangın(Hektar) (N)", "Miktar (Çekim) (O)", "TK-9 gorev(Hektar) (P)", 
      "Uydu (Dk) (Q)", "AÇIKLAMA (R)"
    ];

    const aoaData = [headers];

    listToExport.forEach(log => {
      const isYangin = (log.gorevTipi || log["Görev Tipi"]) === 'Yangın Uçuşu';
      aoaData.push([
        log.id || log["Sıra No"] || "",
        log.tarih || log["Tarih"] || "",
        log.kaptanPilot || log["Kaptan Pilot"] || "",
        log.ikinciPilot || log["2. Pilot"] || log["ikinciPilot"] || "",
        log.teknisyen1 || log["Teknisyen"] || log["Teknisyen 1"] || "",
        log.operator1 || log["Operator"] || log["Operator 1"] || "",
        log.teknisyen2 || log["Teknisyen (2)"] || log["Teknisyen 2"] || "",
        log.operator2 || log["Operator (2)"] || log["Operator 2"] || "",
        log.gorevTipi || log["Görev Tipi"] || "",
        log.gorevBolgesi || log["Görev Bölgesi"] || "",
        log.kalkis || log["Kalkış"] || "",
        log.inis || log["İniş"] || "",
        log.ucusSuresi || log["Uçuş Süresi"] || "",
        isYangin ? (log.k9YanginHektar || log["TK-9 Yangın(Hektar)"] || "0") : "-",
        log.miktarCekim || log["Miktar (Çekim)"] || "0",
        log.tk9GorevHektar || log["TK-9 gorev(Hektar)"] || "0",
        log.uyduDk || log["Uydu (Dk)"] || "0",
        log.aciklama || log["AÇIKLAMA"] || ""
      ]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(aoaData);
    
    // Calculate total minutes
    let totalMinutes = 0;
    listToExport.forEach(log => {
      const duration = log.ucusSuresi || log["Uçuş Süresi"] || '00:00';
      const [h, m] = duration.split(':').map(Number);
      if (!isNaN(h) && !isNaN(m)) totalMinutes += (h * 60 + m);
    });
    const totalHoursString = `${Math.floor(totalMinutes / 60)}:${(totalMinutes % 60).toString().padStart(2, '0')}`;

    // Add summary row
    aoaData.push([]); // spacer
    
    const sumYangin = Math.round(listToExport.reduce((acc, l) => acc + (Number(l.k9YanginHektar) || 0), 0) * 10) / 10;
    const sumCekim = listToExport.reduce((acc, l) => acc + (Number(l.miktarCekim) || 0), 0);
    const sumGorev = Math.round(listToExport.reduce((acc, l) => acc + (Number(l.tk9GorevHektar) || 0), 0) * 10) / 10;
    const sumUydu = listToExport.reduce((acc, l) => acc + parseUyduDkForSum(l.uyduDk), 0);

    const summaryRow = [
      'TOPLAM', // Col 0 (A)
      '',       // Col 1 (B)
      '',       // Col 2 (C)
      '',       // Col 3 (D)
      '',       // Col 4 (E)
      '',       // Col 5 (F)
      '',       // Col 6 (G)
      '',       // Col 7 (H)
      '',       // Col 8 (I)
      '',       // Col 9 (J)
      '',       // Col 10 (K)
      '',       // Col 11 (L)
      totalHoursString, // Col 12 (M)
      sumYangin, // Col 13 (N)
      sumCekim,  // Col 14 (O)
      sumGorev,  // Col 15 (P)
      sumUydu,   // Col 16 (Q)
      ''        // Col 17 (R)
    ];

    aoaData.push(summaryRow);
    const lastDataRowIndex = aoaData.length - 1;
    XLSX.utils.sheet_add_aoa(worksheet, [summaryRow], { origin: `A${lastDataRowIndex + 1}` });

    // Set full range to ensure all 18 columns are included even if empty
    const range = { s: { r: 0, c: 0 }, e: { r: aoaData.length - 1, c: headers.length - 1 } };
    worksheet['!ref'] = XLSX.utils.encode_range(range);

    // Ensure all cells in range exist to apply borders and alignment
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
        if (!worksheet[cell_ref]) {
          worksheet[cell_ref] = { v: "", t: "s" };
        }
      }
    }

    const colsWidth: any[] = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let maxLen = 12;
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
        const cell = worksheet[cell_ref];
        
        const val = cell.v ? cell.v.toString() : '';
        maxLen = Math.max(maxLen, val.length + 2);

        if (!cell.s) cell.s = {};
        
        // Default styling: Alignment and Border
        cell.s.alignment = { horizontal: "center", vertical: "center" };
        cell.s.border = {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        };

        // Header style (Row 0)
        if (R === 0) {
          cell.s.fill = { fgColor: { rgb: "10B981" } }; // Green/Teal
          cell.s.font = { color: { rgb: "FFFFFF" }, bold: true };
        }

        // Highlight matched pilot if not raw export and we are in data rows
        if (!isRawExport && R > 0 && R < lastDataRowIndex - 1) {
          const pilotFilter = filters.pilot;
          if (pilotFilter) {
            if (isPersonMatch(val, pilotFilter)) {
              cell.s.fill = { fgColor: { rgb: "FB923C" } }; // Orange
              cell.s.font = { bold: true };
            }
          }
        }

        // Total row style (last non-empty row)
        if (R === lastDataRowIndex) {
          cell.s.fill = { fgColor: { rgb: "1E293B" } }; // Dark Slate
          cell.s.font = { color: { rgb: "FFFFFF" }, bold: true };
        }
      }
      colsWidth.push({ wch: Math.min(maxLen, 30) });
    }

    worksheet['!cols'] = colsWidth;
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, isRawExport ? "Ham Veri" : "Ucus Kayitlari");
    XLSX.writeFile(workbook, fileName);
  };

  const exportPersonnelToPDF = async () => {
    setIsExportingPDF(true);
    const doc = new jsPDF();
    const timestamp = new Date().toLocaleDateString('tr-TR');

    const toSafePdfText = (text: string) => {
      if (!text) return '';
      return text
        .replace(/Ğ/g, 'G')
        .replace(/ğ/g, 'g')
        .replace(/Ü/g, 'U')
        .replace(/ü/g, 'u')
        .replace(/Ş/g, 'S')
        .replace(/ş/g, 's')
        .replace(/İ/g, 'I')
        .replace(/ı/g, 'i')
        .replace(/Ö/g, 'O')
        .replace(/ö/g, 'o')
        .replace(/Ç/g, 'C')
        .replace(/ç/g, 'c');
    };

    const getBase64Image = async (url: string): Promise<string | null> => {
      if (!url) return null;

      // Try multiple proxy services to bypass Google Drive's lacks of CORS headers
      const proxies = [
        `https://corsproxy.io/?${encodeURIComponent(url)}`,
        `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(url)}`,
        `https://images1-focus-opensocial.googleusercontent.com/gadgets/proxy?container=focus&refresh=2592000&url=${encodeURIComponent(url)}`,
        `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`
      ];

      for (const proxyUrl of proxies) {
        try {
          console.log(`[PDF Preload] Hitting proxy: ${proxyUrl}`);
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 4000); // 4 seconds limit per proxy
          
          const res = await fetch(proxyUrl, { signal: controller.signal });
          clearTimeout(timeoutId);
          
          if (res.ok) {
            const blob = await res.blob();
            if (blob.size > 100) {
              const b64 = await new Promise<string | null>((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result as string);
                reader.onerror = () => resolve(null);
                reader.readAsDataURL(blob);
              });
              if (b64 && (b64.startsWith('data:image/') || b64.length > 500)) {
                console.log(`[PDF Preload] Succeeded with proxy: ${proxyUrl.split('?')[0]}`);
                return b64;
              }
            }
          }
        } catch (err) {
          console.warn(`[PDF Preload] Proxy fail: ${proxyUrl.split('?')[0]}`, err);
        }
      }

      // Final direct fetch fallback
      try {
        console.log(`[PDF Preload] Trying direct fetch (no CORS proxy): ${url}`);
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 2000);
        const res = await fetch(url, { signal: controller.signal });
        clearTimeout(timeoutId);
        if (res.ok) {
          const blob = await res.blob();
          if (blob.size > 100) {
            const b64 = await new Promise<string | null>((resolve) => {
              const reader = new FileReader();
              reader.onloadend = () => resolve(reader.result as string);
              reader.onerror = () => resolve(null);
              reader.readAsDataURL(blob);
            });
            if (b64) {
              console.log("[PDF Preload] Succeeded with direct fetch!");
              return b64;
            }
          }
        }
      } catch (err) {
        console.warn("[PDF Preload] Direct fetch fail:", err);
      }

      return null;
    };

    function fetchWithTimeout<T>(promise: Promise<T>, timeoutMs: number, fallbackValue: T): Promise<T> {
      return Promise.race([
        promise,
        new Promise<T>((resolve) => setTimeout(() => resolve(fallbackValue), timeoutMs))
      ]);
    }

    const filteredPilots = pilotsInResults
      .filter(p => p && normalize(p).includes(normalize(personnelSearch)))
      .sort();
    const filteredTechnical = technicalInResults
      .filter(p => {
        const isSerkan = p === 'SERKAN KEBAPCI' || p === 'SERKAN KEBABCI';
        if (isSerkan) {
          const isBeforeJan2026 = personnelFilters.startDate && personnelFilters.startDate < '2026-01-01';
          return !!isBeforeJan2026;
        }
        return true;
      })
      .filter(p => p && normalize(p).includes(normalize(personnelSearch)))
      .sort();

    const imageCache: Record<string, string> = {};

    try {
      const allPersonnelWithRoles = [
        ...filteredPilots.map(p => ({ name: p, role: 'Pilot' as const })),
        ...filteredTechnical.map(t => ({ name: t, role: 'Teknisyen' as const }))
      ];

      console.log(`[PDF Preload] Pre-loading photos for ${allPersonnelWithRoles.length} personnel...`);
      const promises = allPersonnelWithRoles.map(async ({ name, role }) => {
        const stats = getPersonStats(name, filteredLogsForPersonnel, personnelData, role);
        if (stats.photoUrl) {
          const imgUrl = getDriveThumbnail(stats.photoUrl);
          if (imgUrl) {
            console.log(`[PDF Preload] Preloading photo for ${name}: ${imgUrl}`);
            const b64 = await fetchWithTimeout(getBase64Image(imgUrl), 8000, null); // 8 seconds limit
            if (b64) {
              imageCache[name] = b64;
              console.log(`[PDF Preload] Preloading photo SUCCESS for ${name}`);
            } else {
              console.warn(`[PDF Preload] Preloading photo FAILED (null returned) for ${name}`);
            }
          }
        } else {
          console.log(`[PDF Preload] No photoUrl found for ${name}`);
        }
      });
      await Promise.all(promises);
      console.log(`[PDF Preload] Pre-loading complete. Cached ${Object.keys(imageCache).length} images.`);
    } catch (err) {
      console.error("Error preloading images:", err);
    }

    let pageNum = 1;

    const formatDateTr = (dateStr: string) => {
      if (!dateStr) return '-';
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        return `${parts[2]}.${parts[1]}.${parts[0]}`;
      }
      return dateStr;
    };

    const dateRangeText = `Ucus Donemi: ${formatDateTr(personnelFilters.startDate)} - ${formatDateTr(personnelFilters.endDate)}`;

    const drawHeader = (sectionTitle: string) => {
      // Draw Page Header
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(16, 185, 129); // emerald-500
      doc.text(toSafePdfText('OGM B-360 PERSONEL UCUS RAPORU'), 105, 12, { align: 'center' });
      
      doc.setFontSize(8);
      doc.setTextColor(100, 116, 139); // slate-500
      doc.text(toSafePdfText(`Rapor Tarihi: ${timestamp} | ${dateRangeText} | Sayfa: ${pageNum}`), 105, 17, { align: 'center' });
      
      // Draw Category title
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.setTextColor(234, 179, 8); // amber-500 tint
      doc.text(toSafePdfText(`— ${sectionTitle} —`), 105, 23, { align: 'center' });
      
      // Draw a subtle line underneath the title
      doc.setDrawColor(6, 78, 59);
      doc.setLineWidth(0.3);
      doc.line(14, 25, 196, 25);
    };

    const drawCard = (personName: string, roleType: 'Pilot' | 'Teknisyen', y: number) => {
      const totalHours = calculateTotalHours(filteredLogsForPersonnel, personName);
      const stats = getPersonStats(personName, filteredLogsForPersonnel, personnelData, roleType);

      // Card container fill
      doc.setFillColor(2, 43, 34); // deep forest green
      doc.roundedRect(14, y, 182, 58, 2, 2, 'F');
      
      // Outline border
      doc.setDrawColor(16, 185, 129);
      doc.setLineWidth(0.2);
      doc.roundedRect(14, y, 182, 58, 2, 2, 'S');

      // Draw Avatar or loaded Base64 image
      if (imageCache[personName]) {
        try {
          doc.addImage(imageCache[personName], 'JPEG', 18, y + 4, 14, 14);
          
          // Draw subtle frame border around the dynamic photo
          doc.setDrawColor(16, 185, 129);
          doc.setLineWidth(0.3);
          doc.roundedRect(18, y + 4, 14, 14, 1.5, 1.5, 'S');
        } catch (imgErr) {
          console.error("PDF addImage failed:", imgErr);
          drawInitialsFallback(personName, y);
        }
      } else {
        drawInitialsFallback(personName, y);
      }

      // Name & Title
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.setTextColor(255, 255, 255);
      
      let displayName = personName.toUpperCase();
      if ((normalize(personName) === 'SERKAN KEBAPCI' || normalize(personName) === 'SERKAN KEBABCI') && roleType === 'Teknisyen') {
        displayName += " (UCUS TEKNISYENI)";
      }
      doc.text(toSafePdfText(displayName), 36, y + 9);

      // Title & Location
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7.5);
      doc.setTextColor(52, 211, 153); // emerald 400
      const titleStr = stats.title ? stats.title.toUpperCase() : ('B-360 ' + roleType.toUpperCase());
      doc.text(toSafePdfText(`${roleType.toUpperCase()} | ${titleStr}`), 36, y + 14);

      // Top-Right: UCUS ADEDI label & count
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.5);
      doc.setTextColor(110, 231, 183); // light emerald
      doc.text(toSafePdfText("UCUS ADEDI"), 188, y + 8, { align: 'right' });

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.setTextColor(255, 255, 255);
      doc.text(toSafePdfText(`${stats.count} Ucus`), 188, y + 14, { align: 'right' });

      // Inside Sub-Columns layout:
      // Column 1: "TOPLAM SURE" Box
      doc.setFillColor(1, 33, 26); // even darker green
      doc.roundedRect(18, y + 21, 80, 33, 1, 1, 'F');

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.5);
      doc.setTextColor(110, 231, 183);
      doc.text(toSafePdfText("TOPLAM SURE"), 22, y + 26);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(251, 146, 60); // orange-400
      doc.text(totalHours, 22, y + 32);

      // Stats rows
      doc.setFontSize(6);
      
      // TOP. HEKTAR
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(110, 231, 183);
      doc.text(toSafePdfText("TOP. HEKTAR:"), 22, y + 38);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(52, 211, 153);
      doc.text(toSafePdfText(String(stats.stats.totalHektar)), 94, y + 38, { align: 'right' });

      // YANGIN
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(200, 253, 222);
      doc.text(toSafePdfText("  YANGIN:"), 22, y + 41);
      doc.text(toSafePdfText(String(stats.stats.yanginHektar)), 94, y + 41, { align: 'right' });

      // GOREV
      doc.text(toSafePdfText("  GOREV:"), 22, y + 44);
      doc.text(toSafePdfText(String(stats.stats.gorevHektar)), 94, y + 44, { align: 'right' });

      // UYDU
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(110, 231, 183);
      doc.text(toSafePdfText("UYDU (DK):"), 22, y + 48.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(96, 165, 250); // soft blue
      doc.text(toSafePdfText(String(stats.stats.uyduDk)), 94, y + 48.5, { align: 'right' });

      // CEKIM
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(110, 231, 183);
      doc.text(toSafePdfText("CEKIM:"), 22, y + 51.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(110, 231, 183);
      doc.text(toSafePdfText(String(stats.stats.cekim)), 94, y + 51.5, { align: 'right' });

      // Column 2: "GOREV DAGILIMI" Header
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.5);
      doc.setTextColor(110, 231, 183);
      doc.text(toSafePdfText("GOREV DAGILIMI"), 106, y + 25);

      // Mission type rows:
      const slicedDistribution = stats.distribution.slice(0, 3);
      if (slicedDistribution.length > 0) {
        slicedDistribution.forEach((dist, idx) => {
          const mType = dist[0];
          const mCount = dist[1];
          const rowY = y + 28 + idx * 7.5;
          
          doc.setFillColor(1, 33, 26);
          doc.roundedRect(106, rowY, 82, 6.5, 0.8, 0.8, 'F');

          doc.setFont('helvetica', 'normal');
          doc.setFontSize(6);
          doc.setTextColor(200, 253, 222);
          
          let labelText = mType;
          if (labelText.length > 30) {
            labelText = labelText.substring(0, 28) + '...';
          }
          doc.text(toSafePdfText(labelText), 109, rowY + 4.3);

          doc.setFont('helvetica', 'bold');
          doc.setFontSize(6.5);
          doc.setTextColor(52, 211, 153); // emerald green
          doc.text(toSafePdfText(String(mCount)), 184, rowY + 4.3, { align: 'right' });
        });

        if (stats.distribution.length > 3) {
          doc.setFont('helvetica', 'italic');
          doc.setFontSize(6);
          doc.setTextColor(110, 231, 183);
          doc.text(toSafePdfText(`+ ${stats.distribution.length - 3} tip daha`), 147, y + 53, { align: 'center' });
        }
      } else {
        doc.setFont('helvetica', 'italic');
        doc.setFontSize(6.5);
        doc.setTextColor(110, 231, 183);
        doc.text(toSafePdfText("Ucus verisi bulunamadi."), 147, y + 35, { align: 'center' });
      }
    };

    const drawInitialsFallback = (personName: string, y: number) => {
      // Left Avatar/Photo container block
      doc.setFillColor(6, 78, 59); // lighter dark green
      doc.roundedRect(18, y + 4, 14, 14, 1.5, 1.5, 'F');
      
      // Avatar text (Initials)
      const names = personName.split(' ');
      const initials = names.length >= 2 
        ? (names[0][0] + names[names.length - 1][0]).toUpperCase()
        : personName.substring(0, 2).toUpperCase();
      
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.setTextColor(255, 255, 255);
      doc.text(toSafePdfText(initials), 25, y + 12.5, { align: 'center' });
    };

    let cardIndexOnPage = 0;

    // Process Pilots
    if (filteredPilots.length > 0) {
      drawHeader('UCUS PILOTLARI');
      
      filteredPilots.forEach((pilotName, index) => {
        if (index > 0 && index % 4 === 0) {
          doc.addPage();
          pageNum++;
          cardIndexOnPage = 0;
          drawHeader('UCUS PILOTLARI');
        }
        
        const yCoord = 27 + cardIndexOnPage * 63;
        drawCard(pilotName, 'Pilot', yCoord);
        cardIndexOnPage++;
      });
    }

    // Process Technicians
    if (filteredTechnical.length > 0) {
      // Always start technicians on a new page if we already printed pilots!
      if (filteredPilots.length > 0) {
        doc.addPage();
        pageNum++;
      }
      cardIndexOnPage = 0;
      drawHeader('TEKNIK EKIP (UCUS TEKNISYENLERI)');

      filteredTechnical.forEach((techName, index) => {
        if (index > 0 && index % 4 === 0) {
          doc.addPage();
          pageNum++;
          cardIndexOnPage = 0;
          drawHeader('TEKNIK EKIP (UCUS TEKNISYENLERI)');
        }

        const yCoord = 27 + cardIndexOnPage * 63;
        drawCard(techName, 'Teknisyen', yCoord);
        cardIndexOnPage++;
      });
    }

    if (filteredPilots.length === 0 && filteredTechnical.length === 0) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.setTextColor(16, 185, 129);
      doc.text(toSafePdfText('Arama sonuclarina uygun personel bulunamadi.'), 105, 50, { align: 'center' });
    }

    doc.save(`Personnel_Report_${personnelSearch || 'Tumu'}.pdf`);
    setIsExportingPDF(false);
  };

  const calculateFilteredTotalDuration = () => {
    let totalMinutes = 0;
    filteredLogs.forEach(log => {
      const duration = log.ucusSuresi || '00:00';
      const [h, m] = duration.split(':').map(Number);
      if (!isNaN(h) && !isNaN(m)) {
        totalMinutes += (h * 60 + m);
      }
    });
    const h = Math.floor(totalMinutes / 60);
    const m = totalMinutes % 60;
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
  };

  const shouldHighlight = (value: string) => {
    const v = (value || '').trim();
    if (!v) return false;
    
    // Check pilot filter
    if (filters.pilot) {
      if (isPersonMatch(v, filters.pilot)) return true;
    }
    
    // Check general search term
    if (searchTerm) {
      const s = searchTerm.toLowerCase().trim();
      if (v.toLowerCase().includes(s)) return true;
    }
    
    return false;
  };

  return (
    <div className="min-h-screen bg-forest-base text-emerald-50 font-sans flex flex-col">
      {/* Top Header */}
      <header className="bg-forest-dark border-b border-emerald-800 px-6 py-4 flex flex-row justify-between items-center sticky top-0 z-50">
        <div className="flex items-center gap-3 cursor-pointer" onClick={() => setActiveTab('home')}>
          <div className="w-10 h-10 rounded-full overflow-hidden border-2 border-emerald-500 shadow-lg bg-forest-base flex items-center justify-center">
            <img 
              src="https://cdn.jetphotos.com/full/6/484828_1737509185.jpg" 
              alt="B-360" 
              className="w-full h-full object-cover"
              referrerPolicy="no-referrer"
            />
          </div>
          <div className="hidden sm:block">
            <h1 className="text-sm font-bold tracking-tight uppercase leading-none">Uçuş Kayıt</h1>
            <p className="text-[8px] text-emerald-400 opacity-80 uppercase tracking-widest font-mono font-bold">Veri Tabanı Sistemi</p>
          </div>
        </div>
        
        <div className="flex gap-2">
          {isLoading && <div className="flex items-center px-2"><div className="w-4 h-4 border-2 border-emerald-500 border-t-transparent rounded-full animate-spin"></div></div>}
          {activeTab !== 'home' && (
            <button 
              onClick={() => setActiveTab('home')}
              className="px-4 py-2 rounded text-[10px] font-black bg-emerald-900/40 text-emerald-300 hover:bg-emerald-800 border border-emerald-700 transition-all uppercase"
            >
              ANA SAYFA
            </button>
          )}
          <NavButton active={activeTab === 'personnel'} onClick={() => setActiveTab('personnel')} icon={<Users size={14} />} label="EKİP" />
        </div>
      </header>

      <main className="flex-1 w-full max-w-4xl mx-auto p-4 flex flex-col">
        <AnimatePresence mode="wait">
          {activeTab === 'home' && (
            <motion.div
              key="home"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="flex-1 flex flex-col justify-center items-center gap-6 py-12"
            >
              <h2 className="text-xl font-black uppercase tracking-[0.2em] text-emerald-500/80 mb-4 bg-forest-dark px-6 py-2 rounded-full border border-emerald-800 shadow-xl">
                 OPERASYON MERKEZİ
              </h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 w-full max-w-lg">
                <button 
                  onClick={() => setActiveTab('add')}
                  className="group bg-emerald-600 hover:bg-emerald-500 p-8 rounded-3xl shadow-2xl transition-all active:scale-95 flex flex-col items-center gap-4 text-white border border-emerald-400/30"
                >
                  <div className="w-16 h-16 bg-white/20 rounded-2xl flex items-center justify-center group-hover:scale-110 transition-transform">
                    <PlusCircle size={40} />
                  </div>
                  <div className="text-center">
                    <span className="block text-2xl font-black uppercase tracking-wider">UÇUŞ EKLE</span>
                    <span className="text-xs opacity-70 font-medium">Yeni görev kaydı başlat</span>
                  </div>
                </button>

                <button 
                  onClick={handleOpenKayitlar}
                  className="group bg-forest-dark hover:bg-emerald-900/40 p-8 rounded-3xl shadow-2xl transition-all active:scale-95 flex flex-col items-center gap-4 text-emerald-400 border border-emerald-800"
                >
                  <div className="w-16 h-16 bg-emerald-500/10 rounded-2xl flex items-center justify-center group-hover:scale-110 transition-transform">
                    <Search size={40} />
                  </div>
                  <div className="text-center">
                    <span className="block text-2xl font-black uppercase tracking-wider text-emerald-100">KAYITLAR</span>
                    <span className="text-xs opacity-70 font-medium">Uçuş geçmişini incele</span>
                  </div>
                </button>
              </div>

              <div className="mt-12 text-center text-emerald-900/50 text-[10px] uppercase font-black tracking-[0.5em]">
                B-360 FLIGHT OPERATIONS SYSTEM V2.0
              </div>
            </motion.div>
          )}

          {activeTab === 'add' && (
            <motion.div
              key="add"
              initial={{ opacity: 0, x: -10 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: 10 }}
            >
              {isSuccess ? (
                <div className="flex flex-col items-center justify-center py-20 px-6 bg-forest-dark/80 backdrop-blur-md rounded-3xl border-2 border-emerald-500/50 shadow-[0_0_50px_rgba(16,185,129,0.2)]">
                   <motion.div 
                     initial={{ scale: 0 }}
                     animate={{ scale: 1, rotate: 360 }}
                     transition={{ type: "spring", stiffness: 260, damping: 20 }}
                     className="w-24 h-24 bg-gradient-to-br from-emerald-400 to-emerald-600 rounded-full flex items-center justify-center text-white mb-8 shadow-lg shadow-emerald-500/40"
                   >
                     <CheckCircle2 size={56} />
                   </motion.div>
                   <motion.h2 
                     initial={{ opacity: 0, y: 10 }}
                     animate={{ opacity: 1, y: 0 }}
                     transition={{ delay: 0.2 }}
                     className="text-3xl font-black uppercase tracking-tighter text-white text-center leading-tight"
                   >
                     KAYIT <br/> 
                     <span className="text-emerald-400">TAMAMLANMIŞTIR</span>
                   </motion.h2>
                   <motion.div 
                     initial={{ opacity: 0 }}
                     animate={{ opacity: 1 }}
                     transition={{ delay: 0.5 }}
                     className="flex items-center gap-2 mt-6 px-4 py-2 bg-emerald-500/10 rounded-full border border-emerald-500/20"
                   >
                     <p className="text-emerald-500 font-bold uppercase tracking-[0.2em] text-[9px]">Göklerden Buluta Aktarıldı</p>
                   </motion.div>
                </div>
              ) : (
                <FlightForm onSubmit={addLog} isSubmitting={isSubmitting} personnelData={personnelData} />
              )}
            </motion.div>
          )}

          {activeTab === 'search' && (
            <motion.div
              key="search"
              initial={{ opacity: 0, x: 10 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -10 }}
              className="space-y-6"
            >
              <div className="grid grid-cols-1 md:grid-cols-4 lg:grid-cols-7 gap-4 bg-forest-dark p-6 rounded-lg border border-emerald-800 shadow-xl items-end text-left">
                <div className="md:col-span-1 lg:col-span-2">
                  <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Arama</label>
                  <div className="relative">
                    <Search className={`absolute left-3 top-1/2 -translate-y-1/2 ${isLoading ? 'text-emerald-400 animate-spin' : 'text-emerald-700'}`} size={14} />
                    <input 
                      type="text" 
                      placeholder="Filtrele..."
                      className="w-full bg-forest-base border border-emerald-700 rounded px-3 py-2.5 pl-9 text-xs focus:outline-none focus:border-emerald-500 transition-all font-medium text-emerald-100"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                    />
                  </div>
                </div>
                <div>
                  <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Mürettebat</label>
                  <select 
                    className="w-full bg-forest-base border border-emerald-700 rounded px-3 py-2.5 text-xs focus:outline-none font-medium text-emerald-100"
                    value={filters.pilot}
                    onChange={(e) => setFilters(prev => ({ ...prev, pilot: e.target.value }))}
                  >
                    <option value="">TÜM EKİP</option>
                    <optgroup label="PİLOTLAR" className="bg-forest-base">
                      {pilotsInResults.map(p => <option key={p} value={p}>{getSurname(p)} ({p})</option>)}
                    </optgroup>
                    <optgroup label="TEKNİK / OPS" className="bg-forest-base">
                      {technicalInResults.map(t => <option key={t} value={t}>{getSurname(t)} ({t})</option>)}
                    </optgroup>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Görev Tipi</label>
                  <select 
                    className="w-full bg-forest-base border border-emerald-700 rounded px-3 py-2.5 text-xs focus:outline-none font-medium text-emerald-100"
                    value={filters.tip}
                    onChange={(e) => setFilters(prev => ({ ...prev, tip: e.target.value }))}
                  >
                    <option value="">TÜM GÖREVLER</option>
                    {GOREV_TIPLERI.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Başlangıç</label>
                  <input 
                    type="date"
                    className="w-full bg-forest-base border border-emerald-700 rounded px-2 py-2 text-[10px] focus:outline-none text-emerald-100"
                    value={filters.startDate}
                    onChange={(e) => setFilters(prev => ({ ...prev, startDate: e.target.value }))}
                  />
                </div>
                <div>
                  <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Bitiş</label>
                  <input 
                    type="date"
                    className="w-full bg-forest-base border border-emerald-700 rounded px-2 py-2 text-[10px] focus:outline-none text-emerald-100"
                    value={filters.endDate}
                    onChange={(e) => setFilters(prev => ({ ...prev, endDate: e.target.value }))}
                  />
                </div>
                  <div className="flex flex-col gap-2">
                    <button 
                      onClick={fetchExternalData}
                      disabled={isLoading}
                      className="w-full bg-emerald-900/60 hover:bg-emerald-800 text-emerald-400 rounded-lg px-2 py-2 text-[8px] font-black transition-all flex items-center justify-center gap-1 shadow-lg"
                    >
                      <History size={12} className={isLoading ? 'animate-spin' : ''} />
                      GÜNCELLE
                    </button>
                    <div className="grid grid-cols-2 gap-2">
                      <button 
                        onClick={() => exportToExcel(false)}
                        className="bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg px-2 py-2 text-[8px] font-black transition-all flex items-center justify-center gap-1 shadow-lg"
                      >
                        <Download size={12} />
                        İNDİR
                      </button>
                      <button 
                        onClick={() => exportToExcel(true)}
                        className="bg-blue-600 hover:bg-blue-500 text-white rounded-lg px-2 py-2 text-[8px] font-black transition-all flex items-center justify-center gap-1 shadow-lg"
                      >
                        <History size={12} />
                        HAM VERİ
                      </button>
                    </div>
                    <button 
                      onClick={() => {
                        const cy = new Date().getFullYear();
                        setSearchTerm('');
                        setFilters({
                            pilot: '',
                            bolge: '',
                            tip: '',
                            startDate: `${cy}-01-01`,
                            endDate: `${cy + 1}-01-01`
                        });
                      }}
                      className="w-full bg-forest-base border border-emerald-800 text-emerald-700 hover:text-emerald-500 rounded-lg px-2 py-1 text-[7px] font-black transition-all"
                    >
                      FİLTRELERİ TEMİZLE
                    </button>
                </div>
              </div>

              <div className="flex flex-wrap gap-4 px-2">
                <div className="bg-forest-dark px-4 py-2 rounded-lg border border-emerald-800 flex flex-col min-w-[120px]">
                  <span className="text-[9px] text-emerald-500/50 uppercase font-black tracking-widest text-left">Filtrelenmiş Kayıt</span>
                  <span className="text-sm font-bold text-emerald-100 text-left">{filteredLogs.length} Adet</span>
                </div>
                <div className={`px-4 py-2 rounded-lg border flex flex-col transition-all min-w-[140px] text-left ${filters.pilot ? 'bg-orange-500/10 border-orange-500/50 scale-105 shadow-[0_0_15px_rgba(249,115,22,0.1)]' : 'bg-forest-dark border-emerald-800'}`}>
                  <span className={`text-[9px] uppercase font-black tracking-widest ${filters.pilot ? 'text-orange-500 text-left' : 'text-emerald-500/50 text-left'}`}>
                    {filters.pilot ? `${getSurname(filters.pilot)} Toplam Süre` : 'Filtrelenmiş Toplam Süre'}
                  </span>
                  <span className={`text-sm font-bold font-mono italic ${filters.pilot ? 'text-orange-400 text-left' : 'text-emerald-400 text-left'}`}>
                    {calculateFilteredTotalDuration()}
                  </span>
                </div>
                <div className="bg-forest-dark px-4 py-2 rounded-lg border border-emerald-800 flex flex-col min-w-[140px] text-left">
                  <span className="text-[9px] text-emerald-500/50 uppercase font-black tracking-widest text-left">TOPLAM HEKTAR (Y+G)</span>
                  <span className="text-sm font-bold text-red-100 text-left">
                    {(filteredLogs.reduce((acc, l) => acc + (Number(l.k9YanginHektar) || 0) + (Number(l.tk9GorevHektar) || 0), 0)).toLocaleString('tr-TR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })} <span className="text-[10px] opacity-40">HK</span>
                  </span>
                  <div className="flex gap-2 mt-1 opacity-60 text-[8px] font-bold">
                    <span className="text-red-400">Y: {(filteredLogs.reduce((acc, l) => acc + (Number(l.k9YanginHektar) || 0), 0)).toLocaleString('tr-TR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</span>
                    <span className="text-blue-400">G: {(filteredLogs.reduce((acc, l) => acc + (Number(l.tk9GorevHektar) || 0), 0)).toLocaleString('tr-TR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</span>
                  </div>
                </div>
                <div className="bg-forest-dark px-4 py-2 rounded-lg border border-emerald-800 flex flex-col min-w-[150px] text-left">
                  <span className="text-[9px] text-emerald-500/50 uppercase font-black tracking-widest text-left">FİLTRELENMİŞ UYDU / ÇEKİM</span>
                  <div className="flex flex-col gap-0.5 mt-0.5">
                    <span className="text-xs font-bold text-blue-400 text-left">
                      UYDU: {filteredLogs.reduce((acc, l) => acc + parseUyduDkForSum(l.uyduDk), 0).toLocaleString('tr-TR')} <span className="text-[9px] opacity-70">DK</span>
                    </span>
                    <span className="text-xs font-bold text-emerald-300 text-left">
                      ÇEKİM: {filteredLogs.reduce((acc, l) => acc + (Number(l.miktarCekim) || 0), 0).toLocaleString('tr-TR')} <span className="text-[9px] opacity-70">ADET</span>
                    </span>
                  </div>
                </div>
              </div>

              <div className="border border-emerald-800 rounded-lg overflow-hidden flex flex-col bg-forest-dark shadow-2xl">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left min-w-[3000px] table-fixed">
                    <thead>
                        <tr className="bg-forest-base text-emerald-400 text-[11px] uppercase tracking-widest border-b border-emerald-700">
                          <th className="w-20 px-4 py-5 font-black sticky left-0 z-20 bg-forest-base border-r border-emerald-800 shadow-[2px_0_5px_rgba(0,0,0,0.3)]">Sıra No (A)</th>
                          <th className="w-32 px-4 py-5 font-black">Tarih (B)</th>
                          <th className="w-48 px-4 py-5 font-black">Kaptan Pilot (C)</th>
                          <th className="w-48 px-4 py-5 font-black">2. Pilot (D)</th>
                          <th className="w-48 px-4 py-5 font-black text-blue-300">Teknisyen (E)</th>
                          <th className="w-48 px-4 py-5 font-black text-blue-300">Operator (F)</th>
                          <th className="w-48 px-4 py-5 font-black text-emerald-600">Teknisyen (G)</th>
                          <th className="w-48 px-4 py-5 font-black text-emerald-600">Operator (H)</th>
                          <th className="w-48 px-4 py-5 font-black border-l border-emerald-800/30">Görev Tipi (I)</th>
                          <th className="w-56 px-4 py-5 font-black">Görev Bölgesi (J)</th>
                          <th className="w-24 px-4 py-5 font-black text-center">Kalkış (K)</th>
                          <th className="w-24 px-4 py-5 font-black text-center">İniş (L)</th>
                          <th className="w-28 px-4 py-5 font-black text-center">Uçuş Süresi (M)</th>
                          <th className="w-32 px-4 py-5 font-black text-center text-red-400">TK-9 Yangın(Hektar) (N)</th>
                          <th className="w-32 px-4 py-5 font-black text-center">Miktar (Çekim) (O)</th>
                          <th className="w-32 px-4 py-5 font-black text-center text-blue-400">TK-9 gorev(Hektar) (P)</th>
                          <th className="w-32 px-4 py-5 font-black text-center opacity-70">Uydu (Dk) (Q)</th>
                          <th className="w-80 px-4 py-5 font-black">AÇIKLAMA (R)</th>
                          <th className="w-56 px-4 py-5 font-black text-center text-orange-400">İŞLEMLER</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-emerald-800/40">
                      {filteredLogs.map((log, index) => (
                        <tr key={`${log.tarih}_${log.id}_${index}`} className="hover:bg-emerald-400/5 text-[12px] group transition-colors">
                          <td className="px-4 py-4 whitespace-nowrap sticky left-0 z-10 bg-forest-dark border-r border-emerald-900 group-hover:bg-emerald-950 transition-colors shadow-[2px_0_5px_rgba(0,0,0,0.3)]">
                             <div className="flex flex-col items-center">
                               <span className="text-[10px] font-black w-fit px-2 py-0.5 rounded leading-none mb-1 text-emerald-400 bg-emerald-950/80 border border-emerald-700/50">
                                 {log.id}
                               </span>
                             </div>
                          </td>
                          <td className="px-4 py-4 font-mono font-bold text-emerald-100 whitespace-nowrap">
                            {formatDisplayDate(log.tarih)}
                          </td>
                          <td className={`px-4 py-4 font-bold whitespace-nowrap transition-colors ${shouldHighlight(log.kaptanPilot) ? 'text-orange-400 bg-orange-400/10' : 'text-emerald-50'}`}>{log.kaptanPilot}</td>
                          <td className={`px-4 py-4 font-semibold whitespace-nowrap transition-colors ${shouldHighlight(log.ikinciPilot) ? 'text-orange-400 bg-orange-400/10' : 'opacity-70'}`}>{log.ikinciPilot}</td>
                          <td className={`px-4 py-4 font-medium whitespace-nowrap transition-colors ${shouldHighlight(log.teknisyen1) ? 'text-orange-400 bg-orange-400/10 font-bold' : 'text-blue-200/80'}`}>{log.teknisyen1 || '-'}</td>
                          <td className={`px-4 py-4 italic whitespace-nowrap transition-colors ${shouldHighlight(log.operator1) ? 'text-orange-400 bg-orange-400/10 font-bold' : 'text-blue-100/50'}`}>{log.operator1 || '-'}</td>
                          <td className={`px-4 py-4 whitespace-nowrap transition-colors ${shouldHighlight(log.teknisyen2) ? 'text-orange-400 bg-orange-400/10 font-bold' : 'opacity-30'}`}>{log.teknisyen2 || '-'}</td>
                          <td className={`px-4 py-4 whitespace-nowrap transition-colors ${shouldHighlight(log.operator2) ? 'text-orange-400 bg-orange-400/10 font-bold' : 'opacity-30'}`}>{log.operator2 || '-'}</td>
                          <td className="px-4 py-4 font-black text-emerald-400 uppercase tracking-tighter whitespace-nowrap border-l border-emerald-800/30">{log.gorevTipi}</td>
                          <td className="px-4 py-4 font-medium opacity-80 whitespace-nowrap truncate">{log.gorevBolgesi}</td>
                          <td className="px-4 py-4 font-mono opacity-50 text-center whitespace-nowrap">{log.kalkis}</td>
                          <td className="px-4 py-4 font-mono opacity-50 text-center whitespace-nowrap">{log.inis}</td>
                          <td className="px-4 py-4 text-center whitespace-nowrap">
                            <span className="font-mono font-black text-emerald-400 bg-emerald-950/60 px-2 py-1 rounded border border-emerald-800/50">
                              {log.ucusSuresi}
                            </span>
                          </td>
                          <td className="px-4 py-4 text-center font-black text-red-500 whitespace-nowrap bg-red-500/5">
                             {(log.k9YanginHektar || 0) > 0 ? log.k9YanginHektar.toLocaleString('tr-TR') : '-'}
                          </td>
                          <td className="px-4 py-4 text-center font-bold text-emerald-200 whitespace-nowrap">{log.miktarCekim || '-'}</td>
                          <td className="px-4 py-4 text-center text-blue-400 font-black whitespace-nowrap bg-blue-400/5">
                             {(log.tk9GorevHektar || 0) > 0 ? log.tk9GorevHektar.toLocaleString('tr-TR') : '-'}
                          </td>
                          <td className="px-4 py-4 text-center opacity-60 whitespace-nowrap">{log.uyduDk ? `${log.uyduDk} DK` : '-'}</td>
                          <td className="px-4 py-4 opacity-70 italic leading-relaxed text-[10px] truncate max-w-xs" title={log.aciklama}>
                             {log.aciklama}
                          </td>
                          <td className="px-4 py-4 text-center whitespace-nowrap border-l border-emerald-800/30 bg-emerald-950/20">
                            <div className="flex justify-center gap-2">
                              <button
                                onClick={() => setEditingLog(log)}
                                className="px-2.5 py-1.5 bg-blue-600/80 hover:bg-blue-500 hover:text-white border border-blue-500/30 text-emerald-100 text-[10px] font-black rounded flex items-center gap-1 transition-all active:scale-95 shadow-md shadow-blue-950/50"
                              >
                                DÜZENLE
                              </button>
                              <button
                                onClick={() => deleteLog(log.id)}
                                className="px-2.5 py-1.5 bg-red-600/80 hover:bg-red-500 hover:text-white border border-red-500/30 text-emerald-100 text-[10px] font-black rounded flex items-center gap-1 transition-all active:scale-95 shadow-md shadow-red-950/50"
                              >
                                SİL
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot className="bg-emerald-950/40 border-t-2 border-emerald-500 sticky bottom-0">
                      <tr className="text-emerald-100 font-extrabold uppercase tracking-widest text-[11px]">
                        <td className="px-4 py-4 sticky left-0 z-20 bg-emerald-900 border-r border-emerald-800 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">TOPLAM</td>
                        <td colSpan={11} className="px-4 py-4 text-right opacity-50">FİLTRELENMİŞ VERİLER TOPLAM SÜRE:</td>
                        <td className="px-4 py-4 text-center bg-emerald-900/60 font-mono text-emerald-400 text-xl shadow-inner">
                          {calculateFilteredTotalDuration()}
                        </td>
                        <td className="px-4 py-4 text-center bg-red-900/20 text-red-500 font-black shadow-inner">
                           {filteredLogs.reduce((acc, l) => acc + (Number(l.k9YanginHektar) || 0), 0).toLocaleString('tr-TR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                        </td>
                        <td className="px-4 py-4 text-center text-emerald-300 font-black shadow-inner">
                           {filteredLogs.reduce((acc, l) => acc + (Number(l.miktarCekim) || 0), 0).toLocaleString('tr-TR')}
                        </td>
                        <td className="px-4 py-4 text-center bg-blue-900/20 text-blue-400 font-black shadow-inner">
                           {filteredLogs.reduce((acc, l) => acc + (Number(l.tk9GorevHektar) || 0), 0).toLocaleString('tr-TR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                        </td>
                        <td className="px-4 py-4 text-center bg-emerald-950 text-blue-300 font-black shadow-inner font-mono text-sm leading-none">
                           {filteredLogs.reduce((acc, l) => acc + parseUyduDkForSum(l.uyduDk), 0).toLocaleString('tr-TR')} <span className="text-[9px] opacity-70">DK</span>
                        </td>
                        <td className="px-4 py-4"></td>
                        <td className="px-4 py-4"></td>
                      </tr>
                    </tfoot>
                  </table>
                  {filteredLogs.length === 0 && (
                    <div className="py-20 text-center flex flex-col items-center gap-4">
                      {isLoading ? (
                         <div className="flex flex-col items-center gap-3">
                            <History size={40} className="text-emerald-500 animate-spin opacity-20" />
                            <span className="text-xs font-black text-emerald-400 tracking-widest uppercase">Veriler Çekiliyor...</span>
                         </div>
                      ) : (
                        <>
                          <div className="opacity-30 uppercase font-black tracking-widest text-xs">
                            Kayıt bulunamadı
                          </div>
                          <div className="flex flex-wrap justify-center gap-3">
                             <button 
                               onClick={fetchExternalData}
                               className="px-6 py-3 bg-emerald-900/30 border border-emerald-800 rounded-xl text-emerald-400 font-black text-[10px] hover:bg-emerald-800 transition-all flex items-center gap-2"
                             >
                               <History size={14} /> EXCEL'İ TEKRAR DENE
                             </button>

                             <a 
                               href={GAS_URL}
                               target="_blank"
                               rel="noreferrer"
                               className="px-6 py-3 bg-blue-900/30 border border-blue-800 rounded-xl text-blue-400 font-black text-[10px] hover:bg-blue-800 transition-all flex items-center gap-2"
                             >
                               <ExternalLink size={14} /> VERİ LİNKİNİ AÇ
                             </a>
                          </div>
                        </>
                      )}
                      
                      {syncError && (
                         <div className="max-w-xl mx-auto mt-6 p-6 bg-red-950/20 border border-red-900/50 rounded-2xl text-left bg-gradient-to-br from-red-950/30 to-transparent">
                           <h4 className="text-red-400 font-black text-xs uppercase mb-3 flex items-center gap-2">
                             <AlertCircle size={16} /> BAĞLANTI SORUNU TESPİT EDİLDİ
                           </h4>
                           <ol className="text-[10px] text-red-300 opacity-80 space-y-2 list-decimal pl-4 font-medium leading-relaxed">
                             <li>Yukarıdaki <b>"VERİ LİNKİNİ AÇ"</b> butonuna tıklayın.</li>
                             <li>Açılan beyaz sayfadaki tüm yazıları (JSON kodlarını) seçip kopyalayın (CTRL+A, CTRL+C).</li>
                           </ol>
                           <p className="mt-4 pt-4 border-t border-red-900/30 text-[9px] text-red-400/50 italic font-bold">
                             Not: Sheets tarafındaki "Erişimi Olanlar: Herkes" ayarı yapılırsa bu adımlara gerek kalmaz.
                           </p>
                         </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'personnel' && (
            <motion.div
              key="personnel"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <div className="bg-forest-dark p-6 rounded-2xl border border-emerald-800 shadow-xl flex flex-col xl:flex-row justify-between items-stretch gap-4">
                <div className="flex flex-col md:flex-row gap-4 flex-1 items-stretch">
                  <div className="flex-1 min-w-[200px]">
                     <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Soyisim / İsim ile Sorgula</label>
                     <div className="relative">
                       <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-emerald-700" size={16} />
                       <input 
                         type="text" 
                         placeholder="Aramak istediğiniz personelin soyismini yazın..."
                         className="w-full bg-forest-base border border-emerald-700 rounded px-4 py-3 pl-10 text-sm focus:outline-none focus:border-emerald-500 transition-all font-medium text-emerald-100"
                         value={personnelSearch}
                         onChange={(e) => setPersonnelSearch(e.target.value)}
                       />
                     </div>
                  </div>
                  <div className="w-full md:w-44">
                    <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Başlangıç Tarihi</label>
                    <input 
                      type="date"
                      className="w-full bg-forest-base border border-emerald-700 rounded px-3 py-3 text-xs focus:outline-none focus:border-emerald-500 transition-all text-emerald-100 font-medium"
                      value={personnelFilters.startDate}
                      onChange={(e) => setPersonnelFilters(prev => ({ ...prev, startDate: e.target.value }))}
                    />
                  </div>
                  <div className="w-full md:w-44">
                    <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Bitiş Tarihi</label>
                    <input 
                      type="date"
                      className="w-full bg-forest-base border border-emerald-700 rounded px-3 py-3 text-xs focus:outline-none focus:border-emerald-500 transition-all text-emerald-100 font-medium"
                      value={personnelFilters.endDate}
                      onChange={(e) => setPersonnelFilters(prev => ({ ...prev, endDate: e.target.value }))}
                    />
                  </div>
                  <div className="flex items-end">
                    <button 
                      onClick={() => {
                        const cy = new Date().getFullYear();
                        setPersonnelSearch('');
                        setPersonnelFilters({
                          startDate: `${cy}-01-01`,
                          endDate: `${cy + 1}-01-01`
                        });
                      }}
                      className="h-[46px] w-full md:w-auto px-4 bg-emerald-950/20 hover:bg-emerald-900/40 text-emerald-400 border border-emerald-700 rounded-lg text-[9px] font-black flex items-center justify-center gap-1.5 transition-all text-center uppercase"
                      title="Filtreleri Sıfırla"
                    >
                      SIFIRLA
                    </button>
                  </div>
                </div>
                
                <div className="flex gap-2 items-end">
                  <button 
                    onClick={exportPersonnelToPDF}
                    disabled={isExportingPDF}
                    className="h-[46px] w-full md:w-auto flex items-center justify-center gap-2 px-6 bg-red-900/40 border border-red-700 rounded-xl text-red-200 font-black text-[10px] hover:bg-red-800 transition-all active:scale-95 disabled:opacity-50 shadow-lg"
                  >
                    {isExportingPDF ? (
                      <>
                        <div className="w-3.5 h-3.5 border-2 border-red-200 border-t-transparent rounded-full animate-spin"></div>
                        HAZIRLANIYOR...
                      </>
                    ) : (
                      <>
                        <FileText size={14} />
                        PDF İNDİR
                      </>
                    )}
                  </button>
                  <button 
                    onClick={fetchExternalData}
                    disabled={isLoading}
                    className="h-[46px] w-full md:w-auto flex items-center justify-center gap-2 px-6 bg-emerald-900/40 border border-emerald-700 rounded-xl text-emerald-400 font-black text-[10px] hover:bg-emerald-800 transition-all active:scale-95 disabled:opacity-50"
                  >
                    <History size={14} className={isLoading ? 'animate-spin' : ''} />
                    EXCEL GÜNCELLE
                  </button>
                </div>
              </div>

              {syncError && (
                <div className="bg-red-950/20 border border-red-900/50 p-4 rounded-xl">
                  <p className="text-xs text-red-400 font-bold mb-2 flex items-center gap-2">
                    <History size={14} /> SENKRONİZASYON HATASI
                  </p>
                  <p className="text-[10px] text-red-300 opacity-80 leading-relaxed">
                    {syncError}
                  </p>
                  <a 
                    href={GAS_URL} 
                    target="_blank" 
                    rel="noopener noreferrer"
                    className="mt-3 inline-block text-[10px] font-black underline text-emerald-400 hover:text-emerald-300"
                  >
                    VERİ LİNKİNİ YENİ SEKMEDE AÇ
                  </a>
                </div>
              )}

              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                 <PersonnelPanel 
                    title="Uçuş Pilotları" 
                    data={pilotsInResults.filter(p => normalize(p).includes(normalize(personnelSearch)))} 
                    type="Pilot" 
                    logs={filteredLogsForPersonnel} 
                    personnelData={personnelData}
                 />
                 <PersonnelPanel 
                    title="Teknik Ekip & Operatörler" 
                    data={technicalInResults
                      .filter(p => {
                        const isSerkan = p === 'SERKAN KEBAPCI' || p === 'SERKAN KEBABCI';
                        if (isSerkan) {
                          const isBeforeJan2026 = personnelFilters.startDate && personnelFilters.startDate < '2026-01-01';
                          return !!isBeforeJan2026;
                        }
                        return true;
                      })
                      .filter(p => normalize(p).includes(normalize(personnelSearch)))} 
                    type="Teknisyen" 
                    logs={filteredLogsForPersonnel} 
                    personnelData={personnelData}
                 />
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="mt-auto p-4 bg-forest-dark border-t border-emerald-800 flex flex-col md:flex-row justify-between items-center gap-4">
        <div className="flex gap-8 items-center text-[10px] text-emerald-400 font-black uppercase tracking-wider">
          <span className="flex items-center gap-2"><Database size={14} className="opacity-50" /> TOPLAM UÇUŞ: {logs.length}</span>
          <span className="opacity-20">|</span>
          <span className="flex items-center gap-2">
            <Clock size={14} className="opacity-50" /> 
            SON KAYIT: {(() => {
              if (logs.length === 0) return 'N/A';
              const maxIdLog = logs.reduce((max, log) => {
                const maxId = parseInt(max.id) || 0;
                const logId = parseInt(log.id) || 0;
                return logId > maxId ? log : max;
              }, logs[0]);
              return formatDisplayDate(maxIdLog?.tarih) || 'N/A';
            })()}
          </span>
        </div>
        <p className="text-[10px] opacity-40 font-mono tracking-widest uppercase">© 2026 Hava Araçları Bakım ve Teknik Şube Müdürlüğü | Secure B-360 DB</p>
      </footer>

      {showPasswordModal && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md flex items-center justify-center z-[100] p-4">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95, y: 10 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95, y: 10 }}
            className="bg-forest-dark border-2 border-emerald-500/30 rounded-3xl p-6 sm:p-8 max-w-sm w-full shadow-[0_0_50px_rgba(16,185,129,0.15)] flex flex-col items-center text-center gap-6"
          >
            <div className="w-16 h-16 bg-emerald-500/10 rounded-2xl flex items-center justify-center text-emerald-400 border border-emerald-500/20 shadow-inner">
              <Lock size={32} />
            </div>
            
            <div>
              <h3 className="text-xl font-bold tracking-tight text-emerald-100">GÜVENLİK GEÇİDİ</h3>
              <p className="text-xs text-emerald-400/70 mt-1 uppercase font-semibold tracking-wider">UÇUŞ KAYITLARINI İNCELEMEK İÇİN ŞİFRE GİRİNİZ</p>
            </div>

            <form onSubmit={handleVerifyPassword} className="w-full flex flex-col gap-4">
              <div className="flex flex-col gap-1.5 text-left">
                <input 
                  type="password"
                  placeholder="Şifre Girin"
                  autoFocus
                  className="w-full bg-forest-base/60 text-center text-xl font-mono tracking-widest text-emerald-100 py-3.5 rounded-xl border border-emerald-800 focus:outline-none focus:border-emerald-500 shadow-inner"
                  value={passwordInput}
                  onChange={(e) => {
                    setPasswordInput(e.target.value);
                    setPasswordError('');
                  }}
                />
                {passwordError && (
                  <p className="text-red-400 text-[10px] font-black uppercase tracking-wider text-center mt-1 animate-pulse flex items-center justify-center gap-1">
                    <AlertCircle size={10} /> {passwordError}
                  </p>
                )}
              </div>

              <div className="flex gap-3 mt-2 w-full">
                <button 
                  type="button"
                  onClick={() => setShowPasswordModal(false)}
                  className="flex-1 bg-forest-base/50 text-emerald-400 border border-emerald-950 font-bold py-3 rounded-lg text-xs uppercase tracking-wider hover:bg-emerald-950 transition-all"
                >
                  İPTAL
                </button>
                <button 
                  type="submit"
                  className="flex-1 bg-emerald-600 text-white font-black py-3 rounded-lg text-xs uppercase tracking-wider hover:bg-emerald-500 shadow-lg shadow-emerald-600/20 active:scale-95 transition-all"
                >
                  ONAYLA
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {editingLog && (
        <EditLogModal 
          log={editingLog} 
          onClose={() => setEditingLog(null)} 
          onSave={updateLog}
          isSubmitting={isSubmitting}
          personnelData={personnelData}
        />
      )}
    </div>
  );
}

function EditLogModal({ log, onClose, onSave, isSubmitting, personnelData }: {
  log: FlightLog;
  onClose: () => void;
  onSave: (id: string, updatedFields: Omit<FlightLog, 'id' | 'ucusSuresi'>) => Promise<void>;
  isSubmitting: boolean;
  personnelData: Record<string, PersonData>;
}) {
  const [formData, setFormData] = useState<Omit<FlightLog, 'id' | 'ucusSuresi'>>({
    tarih: log.tarih || '',
    kaptanPilot: log.kaptanPilot || '',
    ikinciPilot: log.ikinciPilot || '',
    teknisyen1: log.teknisyen1 || '',
    operator1: log.operator1 || '',
    teknisyen2: log.teknisyen2 || '',
    operator2: log.operator2 || '',
    gorevTipi: log.gorevTipi || '',
    gorevBolgesi: log.gorevBolgesi || '',
    kalkis: log.kalkis || '',
    inis: log.inis || '',
    k9YanginHektar: Number(log.k9YanginHektar) || 0,
    miktarCekim: Number(log.miktarCekim) || 0,
    tk9GorevHektar: Number(log.tk9GorevHektar) || 0,
    uyduDk: Number(log.uyduDk) || 0,
    aciklama: log.aciklama || ''
  });

  const pilotOptions = Array.from(new Set([
    ...PILOTS, 
    ...Object.values(personnelData).filter(p => !p.role || p.role.includes('Pilot') || p.role.includes('Kaptan')).map(p => p.fullName)
  ])).sort((a, b) => getSurname(a).localeCompare(getSurname(b)));

  const technicalOptions = Array.from(new Set([
    ...TECHNICIANS, 
    ...Object.values(personnelData).filter(p => !p.role || (!p.role.includes('Pilot') && !p.role.includes('Kaptan'))).map(p => p.fullName)
  ])).sort((a, b) => getSurname(a).localeCompare(getSurname(b)));

  const labelStyle = "text-[9px] text-emerald-400 uppercase font-black mb-1 block tracking-widest pl-1 text-left";
  const inputStyle = "w-full bg-forest-base border border-emerald-800 rounded px-3 py-2 text-xs text-emerald-50 focus:outline-none focus:border-emerald-500 placeholder:opacity-25 transition-all font-medium text-left";

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.tarih || !formData.gorevTipi) {
      alert('Tarih ve Görev Tipi alanları zorunludur.');
      return;
    }
    if (!formData.kaptanPilot || !formData.ikinciPilot) {
      alert('Pilotlar (Kaptan ve 2. Pilot) zorunludur.');
      return;
    }
    if (!formData.kalkis || !formData.inis) {
      alert('Kalkış ve İniş saatleri zorunludur.');
      return;
    }
    onSave(log.id, formData);
  };

  return (
    <div className="fixed inset-0 bg-black/80 backdrop-blur-md flex items-center justify-center z-[100] p-4 overflow-y-auto">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 15 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 15 }}
        className="bg-forest-dark border-2 border-emerald-500/30 rounded-3xl p-6 md:p-8 max-w-4xl w-full shadow-[0_0_50px_rgba(16,185,129,0.15)] flex flex-col gap-6 max-h-[90vh] overflow-y-auto"
      >
        <div className="flex justify-between items-center border-b border-emerald-800/60 pb-4">
          <div className="text-left">
            <h3 className="text-lg font-black tracking-tight text-emerald-100 flex items-center gap-2">
              <span className="bg-emerald-950 px-2 py-0.5 rounded text-emerald-400 text-xs border border-emerald-700 font-mono">
                SIRA NO: {log.id}
              </span> 
              KAYIT GÜNCELLEME
            </h3>
            <p className="text-[9px] text-emerald-400/60 uppercase font-bold tracking-wider mt-1">E-Tablo verisiyle gerçek zamanlı senkronize olur.</p>
          </div>
          <button 
            type="button" 
            onClick={onClose}
            className="text-emerald-500 hover:text-emerald-300 font-black text-sm p-1"
          >
            KAPAT
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex flex-col gap-6">
          {/* Section 1: Genel Bilgiler */}
          <div>
            <h4 className="text-[10px] text-emerald-500 font-extrabold uppercase tracking-[0.2em] mb-3 pb-1 border-b border-emerald-800/30 text-left font-sans">
              GENEL BİLGİLER
            </h4>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="text-left">
                <label className={labelStyle}>Tarih</label>
                <input 
                  type="date" 
                  value={formData.tarih} 
                  onChange={e => setFormData({ ...formData, tarih: e.target.value })}
                  className={inputStyle} 
                />
              </div>
              <div className="text-left">
                <label className={labelStyle}>Görev Tipi</label>
                <select 
                  value={formData.gorevTipi} 
                  onChange={e => setFormData({ ...formData, gorevTipi: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">Görev Tipi Seçiniz</option>
                  {GOREV_TIPLERI.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
              <div className="text-left">
                <label className={labelStyle}>Görev Bölgesi</label>
                <input 
                  type="text" 
                  placeholder="Bölge giriniz (örn: Muğla)" 
                  value={formData.gorevBolgesi} 
                  onChange={e => setFormData({ ...formData, gorevBolgesi: e.target.value })}
                  className={inputStyle} 
                />
              </div>
            </div>
          </div>

          {/* Section 2: Uçuş Ekibi */}
          <div>
            <h4 className="text-[10px] text-emerald-500 font-extrabold uppercase tracking-[0.2em] mb-3 pb-1 border-b border-emerald-800/30 text-left font-sans">
              UÇUŞ EKİBİ
            </h4>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="text-left">
                <label className={labelStyle}>Kaptan Pilot</label>
                <select 
                  value={formData.kaptanPilot} 
                  onChange={e => setFormData({ ...formData, kaptanPilot: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">Kaptan Seçiniz</option>
                  {pilotOptions.map(p => (
                    <option key={p} value={p}>{p}</option>
                  ))}
                </select>
              </div>
              <div className="text-left">
                <label className={labelStyle}>2. Pilot</label>
                <select 
                  value={formData.ikinciPilot} 
                  onChange={e => setFormData({ ...formData, ikinciPilot: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">2. Pilot Seçiniz</option>
                  {pilotOptions.map(p => (
                    <option key={p} value={p}>{p}</option>
                  ))}
                </select>
              </div>
              <div className="text-left">
                <label className={labelStyle}>Teknisyen 1 (E)</label>
                <select 
                  value={formData.teknisyen1} 
                  onChange={e => setFormData({ ...formData, teknisyen1: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">Seçiniz</option>
                  {technicalOptions.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
              <div className="text-left">
                <label className={labelStyle}>Operatör 1 (F)</label>
                <select 
                  value={formData.operator1} 
                  onChange={e => setFormData({ ...formData, operator1: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">Seçiniz</option>
                  {technicalOptions.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
              <div className="text-left">
                <label className={labelStyle}>Teknisyen 2 (G)</label>
                <select 
                  value={formData.teknisyen2} 
                  onChange={e => setFormData({ ...formData, teknisyen2: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">Seçiniz (İsteğe Bağlı)</option>
                  {technicalOptions.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
              <div className="text-left">
                <label className={labelStyle}>Operatör 2 (H)</label>
                <select 
                  value={formData.operator2} 
                  onChange={e => setFormData({ ...formData, operator2: e.target.value })}
                  className={inputStyle}
                >
                  <option value="">Seçiniz (İsteğe Bağlı)</option>
                  {technicalOptions.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Section 3: Uçuş Detayları */}
          <div>
            <h4 className="text-[10px] text-emerald-500 font-extrabold uppercase tracking-[0.2em] mb-3 pb-1 border-b border-emerald-800/30 text-left font-sans">
              UÇUŞ DETAYLARI
            </h4>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div className="text-left">
                <label className={labelStyle}>Kalkış</label>
                <input 
                  type="time" 
                  value={formData.kalkis} 
                  onChange={e => setFormData({ ...formData, kalkis: e.target.value })}
                  className={inputStyle} 
                />
              </div>
              <div className="text-left">
                <label className={labelStyle}>İniş</label>
                <input 
                  type="time" 
                  value={formData.inis} 
                  onChange={e => setFormData({ ...formData, inis: e.target.value })}
                  className={inputStyle} 
                />
              </div>
              <div className="text-left">
                <label className={labelStyle}>Uçuş Süresi (M)</label>
                <div className="w-full bg-forest-base/40 border border-emerald-800/50 rounded px-3 py-2 text-xs text-emerald-300 font-mono font-bold select-none h-[34px] flex items-center justify-center">
                  {formData.kalkis && formData.inis ? calculateDuration(formData.kalkis, formData.inis) : '00:00'}
                </div>
              </div>
              <div className="text-left">
                <label className={labelStyle}>Uydu (Dk) (Q)</label>
                <input 
                  type="number" 
                  value={formData.uyduDk} 
                  onChange={e => setFormData({ ...formData, uyduDk: Number(e.target.value) || 0 })}
                  className={inputStyle} 
                />
              </div>
              <div className="text-left">
                <label className={labelStyle}>K9 Yangın (Hektar) (N)</label>
                <input 
                  type="number" 
                  step="0.1"
                  value={formData.k9YanginHektar} 
                  onChange={e => setFormData({ ...formData, k9YanginHektar: Number(e.target.value) || 0 })}
                  disabled={formData.gorevTipi !== 'Yangın Uçuşu'}
                  className={`${inputStyle} disabled:opacity-30`} 
                />
              </div>
              <div className="text-left">
                <label className={labelStyle}>Miktar (Çekim) (O)</label>
                <input 
                  type="number" 
                  value={formData.miktarCekim} 
                  onChange={e => setFormData({ ...formData, miktarCekim: Number(e.target.value) || 0 })}
                  className={inputStyle} 
                />
              </div>
              <div className="text-left col-span-2">
                <label className={labelStyle}>TK9 Görev Hektar (P)</label>
                <input 
                  type="number" 
                  step="0.1"
                  value={formData.tk9GorevHektar} 
                  onChange={e => setFormData({ ...formData, tk9GorevHektar: Number(e.target.value) || 0 })}
                  className={inputStyle} 
                />
              </div>
            </div>
            
            <div className="text-left mt-4">
              <label className={labelStyle}>Açıklama (R)</label>
              <textarea 
                rows={2}
                placeholder="Açıklama giriniz..."
                value={formData.aciklama} 
                onChange={e => setFormData({ ...formData, aciklama: e.target.value })}
                className={`${inputStyle} resize-none`} 
              />
            </div>
          </div>

          {/* Form Actions */}
          <div className="flex gap-4 border-t border-emerald-800/40 pt-6">
            <button 
              type="button" 
              onClick={onClose}
              className="flex-1 bg-forest-base border border-emerald-950 text-emerald-400 font-bold py-3.5 rounded-xl text-xs uppercase tracking-wider hover:bg-emerald-950/60 transition-all text-center"
            >
              Vazgeç
            </button>
            <button 
              type="submit" 
              disabled={isSubmitting}
              className="flex-1 bg-gradient-to-r from-emerald-600 to-emerald-700 text-white font-black py-3.5 rounded-xl text-xs uppercase tracking-wider hover:from-emerald-500 hover:to-emerald-600 shadow-lg shadow-emerald-900/40 active:scale-95 transition-all disabled:opacity-50 text-center"
            >
              {isSubmitting ? 'Güncelleniyor...' : 'Değişiklikleri Kaydet'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function PersonnelPanel({ title, data, type, logs, personnelData }: { title: string, data: string[], type: string, logs: FlightLog[], personnelData: Record<string, PersonData> }) {
  return (
    <div className="bg-forest-dark border border-emerald-800 rounded-xl overflow-hidden p-6 shadow-xl h-full">
      <div className="flex items-center gap-2 mb-6 border-b border-emerald-800/50 pb-4">
        <div className="w-2 h-2 bg-emerald-500 rounded-full"></div>
        <h2 className="text-xs font-black uppercase tracking-widest text-emerald-400">{title}</h2>
      </div>
      <div className="grid gap-4">
        {data.length > 0 ? data.map((name, i) => {
          const totalHours = calculateTotalHours(logs, name);
          const stats = getPersonStats(name, logs, personnelData, type);
          return (
            <div key={i} className="bg-forest-base/50 rounded-xl border border-emerald-900/30 group hover:border-emerald-500/50 transition-all overflow-hidden flex flex-col">
              <div className="flex items-center justify-between p-4 bg-emerald-950/20">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 rounded-lg bg-forest-dark flex items-center justify-center text-emerald-600 group-hover:text-emerald-400 transition-colors border border-emerald-800 overflow-hidden shrink-0 shadow-inner">
                      {stats.photoUrl ? (
                         <img 
                           src={getDriveThumbnail(stats.photoUrl)} 
                           alt={name} 
                           referrerPolicy="no-referrer"
                           className="w-full h-full object-cover" 
                         />
                       ) : (
                       stats.role === 'Pilot' ? <Users size={24} /> : stats.role === 'Operatör' ? <Camera size={24} /> : <Database size={24} />
                     )}
                  </div>
                  <div>
                    <h3 className="font-bold text-sm tracking-wide uppercase italic leading-tight text-white">
                      {name}
                      {((normalize(name) === 'SERKAN KEBAPCI' || normalize(name) === 'SERKAN KEBABCI') && type === 'Teknisyen') && (
                        <span className="text-emerald-400 font-bold normal-case text-xs inline-block ml-1 bg-emerald-950/40 px-1.5 py-0.5 border border-emerald-800/60 rounded">
                          (UÇUŞ TEKNİSYENİ)
                        </span>
                      )}
                    </h3>
                    <p className="text-[9px] text-emerald-500 font-black uppercase tracking-widest mt-0.5">
                      {stats.role} {stats.title ? `| ${stats.title}` : ''}
                    </p>
                  </div>
                </div>
                <div className="text-right">
                  <span className="text-[8px] block opacity-40 uppercase font-black mb-0.5">Uçuş Adedi</span>
                  <span className="text-emerald-100 font-black text-xs">{stats.count} Uçuş</span>
                </div>
              </div>

              <div className="p-4 grid grid-cols-2 gap-4 border-t border-emerald-900/30">
                <div className="bg-forest-dark/50 p-3 rounded-lg border border-emerald-800/50 flex flex-col justify-between">
                  <div>
                    <span className="text-[8px] block opacity-40 uppercase font-black mb-1">Toplam Süre</span>
                    <span className="text-orange-400 font-mono font-black text-lg">{totalHours}</span>
                  </div>
                  <div className="mt-2 pt-2 border-t border-emerald-900/20 space-y-1">
                    <div className="flex justify-between text-[9px]">
                      <span className="opacity-40">TOP. HEKTAR:</span>
                      <span className="font-bold text-emerald-400">{stats.stats.totalHektar}</span>
                    </div>
                    <div className="flex justify-between text-[8px] opacity-60">
                      <span>YANGIN:</span>
                      <span>{stats.stats.yanginHektar}</span>
                    </div>
                    <div className="flex justify-between text-[8px] opacity-60">
                      <span>GÖREV:</span>
                      <span>{stats.stats.gorevHektar}</span>
                    </div>
                    <div className="flex justify-between text-[9px] pt-1 mt-1 border-t border-emerald-900/10">
                      <span className="opacity-40">UYDU (DK):</span>
                      <span className="font-bold text-blue-400">{stats.stats.uyduDk}</span>
                    </div>
                    <div className="flex justify-between text-[9px]">
                      <span className="opacity-40">ÇEKİM:</span>
                      <span className="font-bold text-emerald-300">{stats.stats.cekim}</span>
                    </div>
                  </div>
                </div>
                <div className="flex flex-col gap-1.5 overflow-hidden">
                   <span className="text-[8px] block opacity-40 uppercase font-black mb-1">Görev Dağılımı</span>
                   <div className="flex flex-col gap-1">
                      {stats.distribution.slice(0, 3).map(([mType, mCount]) => (
                        <div key={mType} className="flex justify-between items-center bg-forest-dark/30 px-2 py-1 rounded text-[9px] border border-emerald-900/20">
                          <span className="opacity-60 truncate mr-2">{mType}</span>
                          <span className="font-bold text-emerald-400">{mCount}</span>
                        </div>
                      ))}
                      {stats.distribution.length > 3 && (
                        <span className="text-[8px] opacity-30 text-center italic mt-1">+ {stats.distribution.length - 3} tip</span>
                      )}
                      {stats.distribution.length === 0 && (
                        <span className="text-[8px] opacity-20 italic">Veri yok</span>
                      )}
                   </div>
                </div>
              </div>
            </div>
          );
        }) : (
          <p className="text-center py-8 text-[10px] uppercase font-bold text-emerald-900/50 italic tracking-widest">Arama kritiğine uygun personel bulunamadı.</p>
        )}
      </div>
    </div>
  );
}

function NavButton({ active, onClick, icon, label }: { active: boolean, onClick: () => void, icon: React.ReactNode, label: string }) {
  return (
    <button 
      onClick={onClick}
      className={`px-6 py-2 rounded text-[11px] font-black flex items-center gap-2 transition-all ${
        active 
          ? 'bg-emerald-600 text-white shadow-lg' 
          : 'text-emerald-500/40 hover:text-emerald-300 hover:bg-forest-dark'
      }`}
    >
      {icon}
      {label}
    </button>
  );
}

function FlightForm({ onSubmit, isSubmitting, personnelData }: { 
  onSubmit: (log: Omit<FlightLog, 'id' | 'ucusSuresi'>) => void, 
  isSubmitting?: boolean,
  personnelData: Record<string, PersonData>
}) {
  const [step, setStep] = useState(1);
  const [formData, setFormData] = useState<Omit<FlightLog, 'id' | 'ucusSuresi'>>({
    tarih: new Date().toISOString().split('T')[0],
    kaptanPilot: '',
    ikinciPilot: '',
    teknisyen1: '',
    operator1: '',
    teknisyen2: '',
    operator2: '',
    gorevTipi: '',
    gorevBolgesi: '',
    kalkis: '',
    inis: '',
    k9YanginHektar: 0,
    miktarCekim: 0,
    tk9GorevHektar: 0,
    uyduDk: 0,
    aciklama: ''
  });

  const nextStep = () => {
    if (step === 1) {
      if (!formData.tarih || !formData.gorevTipi) {
        alert('Tarih ve Görev Tipi zorunludur.');
        return;
      }
      // Kamera Testi ise Mürettebat adımını atla
      if (formData.gorevTipi === 'Kamera Testi') {
        setStep(3);
        return;
      }
    }
    if (step === 2) {
      if (!formData.kaptanPilot || !formData.ikinciPilot) {
        alert('Pilotlar (Kaptan ve 2. Pilot) zorunludur.');
        return;
      }
    }
    setStep(step + 1);
  };

  const prevStep = () => {
    if (step === 3 && formData.gorevTipi === 'Kamera Testi') {
      setStep(1);
      return;
    }
    setStep(step - 1);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.kalkis || !formData.inis) {
      alert('Kalkış ve İniş saatleri zorunludur.');
      return;
    }
    onSubmit(formData);
  };

  const labelStyle = "text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1";
  const inputStyle = "w-full bg-forest-dark border border-emerald-800 rounded px-4 py-4 text-sm text-emerald-50 focus:outline-none focus:border-emerald-500 placeholder:opacity-20 transition-all font-medium";
  
  // Selected names to filter out of other dropdowns
  const selectedNames = [
    formData.kaptanPilot,
    formData.ikinciPilot,
    formData.teknisyen1,
    formData.teknisyen2,
    formData.operator1,
    formData.operator2
  ].filter(Boolean);

  const pilotOptions = Array.from(new Set([
    ...PILOTS, 
    ...Object.values(personnelData).filter(p => !p.role || p.role.includes('Pilot') || p.role.includes('Kaptan')).map(p => p.fullName)
  ])).sort((a, b) => getSurname(a).localeCompare(getSurname(b)));

  const technicalOptions = Array.from(new Set([
    ...TECHNICIANS, 
    ...Object.values(personnelData).filter(p => !p.role || (!p.role.includes('Pilot') && !p.role.includes('Kaptan'))).map(p => p.fullName)
  ])).sort((a, b) => getSurname(a).localeCompare(getSurname(b)));

  return (
    <div className="max-w-md mx-auto w-full">
      {/* Progress Bar */}
      <div className="flex justify-between items-center mb-8 px-2">
        {[1, 2, 3, 4].map((s) => (
          <div key={s} className="flex items-center">
            <div className={`w-8 h-8 rounded-full flex items-center justify-center font-black text-xs border-2 transition-all ${
              step >= s ? 'bg-emerald-600 border-emerald-400 text-white' : 'bg-forest-dark border-emerald-900 text-emerald-900'
            }`}>
              {s}
            </div>
            {s < 4 && <div className={`w-8 md:w-16 h-0.5 mx-1 rounded ${step > s ? 'bg-emerald-500' : 'bg-emerald-900'}`}></div>}
          </div>
        ))}
      </div>

      <motion.form 
        key={step}
        initial={{ opacity: 0, x: 20 }}
        animate={{ opacity: 1, x: 0 }}
        onSubmit={handleSubmit} 
        className="space-y-6"
      >
        {step === 1 && (
          <div className="bg-forest-dark p-6 rounded-2xl border border-emerald-800 shadow-2xl space-y-6 animate-in slide-in-from-right duration-300">
            <h3 className="text-xs font-black uppercase tracking-widest text-emerald-500 border-b border-emerald-800 pb-4">Adım 1: Görev Başlatma</h3>
            <div>
              <label className={labelStyle}>Uçuş Tarihi</label>
              <input type="date" className={inputStyle} value={formData.tarih} onChange={e => setFormData(p => ({...p, tarih: e.target.value}))} />
            </div>
            <div>
              <label className={labelStyle}>Görev Tipi</label>
              <select className={inputStyle} value={formData.gorevTipi} onChange={e => setFormData(p => ({...p, gorevTipi: e.target.value}))}>
                <option value="">SEÇİNİZ...</option>
                {GOREV_TIPLERI.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
          </div>
        )}

        {step === 2 && (
          <div className="bg-forest-dark p-6 rounded-2xl border border-emerald-800 shadow-2xl space-y-6 animate-in slide-in-from-right duration-300">
            <h3 className="text-xs font-black uppercase tracking-widest text-emerald-500 border-b border-emerald-800 pb-4">Adım 2: Mürettebat</h3>
            <div className="grid grid-cols-2 gap-4">
                  <div className="col-span-1">
                    <label className={labelStyle}>Kaptan Pilot</label>
                    <select className={inputStyle} value={formData.kaptanPilot} onChange={e => setFormData(p => ({...p, kaptanPilot: e.target.value}))}>
                       <option value="">PİLOT...</option>
                       {pilotOptions.filter(p => !selectedNames.includes(p) || p === formData.kaptanPilot).map(p => (
                         <option key={p} value={p}>{getSurname(p)} ({p})</option>
                       ))}
                    </select>
                  </div>
                  <div className="col-span-1">
                    <label className={labelStyle}>2. Pilot</label>
                    <select className={inputStyle} value={formData.ikinciPilot} onChange={e => setFormData(p => ({...p, ikinciPilot: e.target.value}))}>
                       <option value="">PİLOT...</option>
                       {pilotOptions.filter(p => !selectedNames.includes(p) || p === formData.ikinciPilot).map(p => (
                         <option key={p} value={p}>{getSurname(p)} ({p})</option>
                       ))}
                    </select>
                  </div>
                </div>
                <div className="space-y-4">
                  <label className={labelStyle}>Teknik Ekip (Opsiyonel)</label>
                  <div className="grid grid-cols-2 gap-2">
                    <select className={inputStyle} value={formData.teknisyen1} onChange={e => setFormData(p => ({...p, teknisyen1: e.target.value}))}>
                      <option value="">TEKNİSYEN 1</option>
                      {technicalOptions.filter(t => !selectedNames.includes(t) || t === formData.teknisyen1).map(t => (
                        <option key={t} value={t}>{getSurname(t)} ({t})</option>
                      ))}
                    </select>
                    <select className={inputStyle} value={formData.teknisyen2} onChange={e => setFormData(p => ({...p, teknisyen2: e.target.value}))}>
                      <option value="">TEKNİSYEN 2</option>
                      {technicalOptions.filter(t => !selectedNames.includes(t) || t === formData.teknisyen2).map(t => (
                        <option key={t} value={t}>{getSurname(t)} ({t})</option>
                      ))}
                    </select>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <select className={inputStyle} value={formData.operator1} onChange={e => setFormData(p => ({...p, operator1: e.target.value}))}>
                      <option value="">OPERATÖR 1</option>
                      {technicalOptions.filter(t => !selectedNames.includes(t) || t === formData.operator1).map(t => (
                        <option key={t} value={t}>{getSurname(t)} ({t})</option>
                      ))}
                    </select>
                    <select className={inputStyle} value={formData.operator2} onChange={e => setFormData(p => ({...p, operator2: e.target.value}))}>
                      <option value="">OPERATÖR 2</option>
                      {technicalOptions.filter(t => !selectedNames.includes(t) || t === formData.operator2).map(t => (
                        <option key={t} value={t}>{getSurname(t)} ({t})</option>
                      ))}
                    </select>
                  </div>
                </div>
          </div>
        )}

        {step === 3 && (
          <div className="bg-forest-dark p-6 rounded-2xl border border-emerald-800 shadow-2xl space-y-6 animate-in slide-in-from-right duration-300">
            <h3 className="text-xs font-black uppercase tracking-widest text-emerald-500 border-b border-emerald-800 pb-4">Adım 3: Mevki & Saat</h3>
            <div>
              <label className={labelStyle}>Görev Bölgesi</label>
              <div className="relative">
                 <MapPin className="absolute left-4 top-1/2 -translate-y-1/2 text-emerald-900" size={16} />
                 <input type="text" placeholder="Bölge..." className={`${inputStyle} pl-12`} value={formData.gorevBolgesi} onChange={e => setFormData(p => ({...p, gorevBolgesi: e.target.value}))} />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className={labelStyle}>Kalkış Saati</label>
                <input type="time" className={`${inputStyle} text-center`} value={formData.kalkis} onChange={e => setFormData(p => ({...p, kalkis: e.target.value}))} />
              </div>
              <div>
                <label className={labelStyle}>İniş Saati</label>
                <input type="time" className={`${inputStyle} text-center`} value={formData.inis} onChange={e => setFormData(p => ({...p, inis: e.target.value}))} />
              </div>
            </div>
            <div className="p-4 bg-emerald-500/5 rounded border border-emerald-500/20 text-center">
               <p className="text-[10px] font-bold text-emerald-500 uppercase tracking-widest mb-1">Durasyon</p>
               <p className="text-3xl font-mono font-black text-emerald-400">{calculateDuration(formData.kalkis, formData.inis)}</p>
            </div>
          </div>
        )}

        {step === 4 && (
          <div className="bg-forest-dark p-6 rounded-2xl border border-emerald-800 shadow-2xl space-y-6 animate-in slide-in-from-right duration-300">
            <h3 className="text-xs font-black uppercase tracking-widest text-emerald-500 border-b border-emerald-800 pb-4">Adım 4: Operasyonel Veriler</h3>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className={labelStyle}>TK-9 Yangın (Hk)</label>
                <div className="relative">
                  <Flame className="absolute left-3 top-1/2 -translate-y-1/2 text-red-900/50" size={14} />
                  <input 
                    type="number" 
                    step="0.01" 
                    min="0"
                    placeholder="0"
                    className={`${inputStyle} pl-10`} 
                    value={formData.k9YanginHektar || ''} 
                    onChange={e => setFormData(p => ({...p, k9YanginHektar: parseFloat(e.target.value) || 0}))} 
                  />
                </div>
              </div>
              <div>
                <label className={labelStyle}>Miktar (Çekim)</label>
                <div className="relative">
                  <Camera className="absolute left-3 top-1/2 -translate-y-1/2 text-emerald-900/50" size={14} />
                  <input 
                    type="number" 
                    min="0"
                    placeholder="0"
                    className={`${inputStyle} pl-10`} 
                    value={formData.miktarCekim || ''} 
                    onChange={e => setFormData(p => ({...p, miktarCekim: parseInt(e.target.value) || 0}))} 
                  />
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className={formData.gorevTipi === 'Yangın Uçuşu' ? 'opacity-30 pointer-events-none' : ''}>
                <label className={labelStyle}>TK-9 Görev (Hk)</label>
                <div className="relative">
                  <ArrowRight className="absolute left-3 top-1/2 -translate-y-1/2 text-blue-900/50" size={14} />
                  <input 
                    type="number" 
                    step="0.01" 
                    min="0"
                    disabled={formData.gorevTipi === 'Yangın Uçuşu'}
                    placeholder="0"
                    className={`${inputStyle} pl-10`} 
                    value={formData.tk9GorevHektar || ''} 
                    onChange={e => setFormData(p => ({...p, tk9GorevHektar: parseFloat(e.target.value) || 0}))} 
                  />
                </div>
              </div>
              <div>
                <label className={labelStyle}>Uydu (Dk)</label>
                <div className="relative">
                  <Satellite className="absolute left-3 top-1/2 -translate-y-1/2 text-blue-900/50" size={14} />
                  <input 
                    type="number" 
                    min="0"
                    placeholder="0"
                    className={`${inputStyle} pl-10`} 
                    value={formData.uyduDk || ''} 
                    onChange={e => setFormData(p => ({...p, uyduDk: parseInt(e.target.value) || 0}))} 
                  />
                </div>
              </div>
            </div>

            <div>
              <label className={labelStyle}>Açıklama</label>
              <textarea 
                className={inputStyle} 
                rows={2} 
                placeholder="Görev detayları..." 
                value={formData.aciklama} 
                onChange={e => setFormData(p => ({...p, aciklama: e.target.value}))} 
              />
            </div>
          </div>
        )}

        <div className="flex gap-4 pt-4">
          {step > 1 && (
            <button 
              type="button" 
              onClick={prevStep}
              className="flex-1 bg-forest-dark text-emerald-400 font-black py-4 rounded-xl border border-emerald-800 uppercase tracking-widest text-xs active:scale-95 transition-all"
            >
              GERİ
            </button>
          )}
          {step < 4 ? (
            <button 
              type="button" 
              onClick={nextStep}
              className="flex-1 bg-emerald-600 text-white font-black py-4 rounded-xl shadow-lg shadow-emerald-600/20 uppercase tracking-widest text-xs active:scale-95 transition-all"
            >
              İLERİ
            </button>
          ) : (
            <button 
              type="submit" 
              disabled={isSubmitting}
              className="flex-1 bg-emerald-500 text-forest-base font-black py-4 rounded-xl shadow-lg shadow-emerald-500/30 uppercase tracking-widest text-xs active:scale-95 transition-all disabled:opacity-50"
            >
              {isSubmitting ? 'GÖNDERİLİYOR...' : 'KAYIT TAMAMLA'}
            </button>
          )}
        </div>
      </motion.form>
    </div>
  );
}

