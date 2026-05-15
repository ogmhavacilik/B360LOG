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
  Trash2
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import { motion, AnimatePresence } from 'motion/react';
import XLSX from 'xlsx-js-style';
import { FlightLog, PILOTS, TECHNICIANS, GOREV_TIPLERI } from './types';

// Google Apps Script URLs
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwpqbW9sHqUZIUwGf4naJN_ZH0aCXvGEYSdqBN6mftyFVZGAofjrnMPfutcO5maBc4/exec';
const PILOT_DATA_URL = 'https://script.google.com/macros/s/AKfycbx5jJmZ6sU8qpSwCFwp41z_7fJYy-buB4BD5686jAoM3xqVw39m3q3iDbkVbQCDUZ5U/exec';
const TECH_DATA_URL = 'https://script.google.com/macros/s/AKfycbzjgsoveqIQS4iLXDBYJrIuYX5yryjj1AKVdejNPCrzL1lSx0EYAEX2ZhcE94uDLpuc/exec';

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
  if (!logCellValue || !selectedFullName) return false;
  
  const rawCell = logCellValue.toString().toUpperCase();
  const rawTarget = selectedFullName.toUpperCase();

  // 1. Literal whole string match (most reliable)
  if (rawCell.includes(rawTarget) || rawTarget.includes(rawCell)) return true;

  const ln = normalize(logCellValue.toString());
  const tn = normalize(selectedFullName);
  if (!ln || !tn) return false;

  const cellParts = ln.split(/\s+/).filter(Boolean);
  const targetParts = tn.split(/\s+/).filter(Boolean);

  // 2. Direct inclusion of normalized strings
  if (ln.includes(tn) || tn.includes(ln)) return true;

  // 3. Cross-part match (any part of name matches part of cell)
  if (targetParts.some(part => part.length >= 3 && cellParts.some(cPart => cPart.includes(part) || part.includes(cPart)))) return true;

  // 4. Surname specific match
  const targetSurname = getSurname(selectedFullName);
  if (targetSurname && targetSurname.length >= 3 && cellParts.includes(targetSurname)) return true;

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
  
  // Standardize Turkish/European format: 1.234,56 -> 1234.56
  if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  } else if (str.match(/^\d+\.\d{3}$/)) {
    // Case like "44.285" (thousands separator)
    str = str.replace(/\./g, '');
  } 
  
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
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
      totalUyduDk += Number(log.uyduDk || 0);
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
  const [personnelData, setPersonnelData] = useState<Record<string, PersonData>>({});
  const [isLoading, setIsLoading] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isSuccess, setIsSuccess] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [personnelSearch, setPersonnelSearch] = useState('');
  const [filters, setFilters] = useState({
    pilot: '',
    bolge: '',
    tip: '',
    startDate: '',
    endDate: ''
  });

  const pilotsInResults = useMemo(() => PILOTS, []);

  const technicalInResults = useMemo(() => TECHNICIANS, []);

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

      const [pilotRes, techRes] = await Promise.all([
        getData(PILOT_DATA_URL),
        getData(TECH_DATA_URL)
      ]);

      const dataMap: Record<string, PersonData> = {};
      
      const processItem = (p: any, type: 'pilot' | 'tech') => {
        if (!p || (Array.isArray(p) && p.length < 2)) return;
        
        let name = '';
        let photo = '';
        let role = '';
        let title = '';

        if (Array.isArray(p)) {
          // Detect headers and skip
          const firstVal = p[0]?.toString().toUpperCase();
          if (firstVal === 'ID' || firstVal === 'AD SOYAD' || p.some(v => v?.toString().toUpperCase() === 'FOTOĞRAF')) return;

          // Search for URL (photo)
          photo = p.find(val => typeof val === 'string' && (val.includes('drive.google.com') || val.includes('http'))) || '';
          
          // Name search: Usually the first long string that isn't a URL or a small specific keyword
          name = p.find((val, idx) => {
            if (typeof val !== 'string' || val.length < 3 || val.includes('http')) return false;
            const up = val.toUpperCase();
            if (['PİLOT', 'TEKNİSYEN', 'KAPTAN', 'OPERATÖR'].includes(up)) return false;
            return true;
          }) || '';
          
          role = p.find(val => {
            if (typeof val !== 'string') return false;
            const up = val.toUpperCase();
            return ['PİLOT', 'TEKNİSYEN', 'OPERATÖR'].some(r => up.includes(r));
          }) || '';

          title = p.find(val => {
             if (typeof val !== 'string' || val.length < 3 || val.includes('http')) return false;
             const up = val.toUpperCase();
             if (['KAPTAN', 'SİSTEM', 'TEKNİK'].some(kw => up.includes(kw))) return true;
             return false;
          }) || '';
        } else {
          name = p.FULL_NAME || p.fullName || p.adSoyad || p.Name || p["AD SOYAD"] || '';
          photo = p.PHOTO_URL || p.photoUrl || p.fotograf || p.Photo || p["FOTOĞRAF"] || p["FOTO"] || '';
          role = p.ROLE || p.role || p["GÖREV"] || '';
          title = p.TITLE || p.title || p["ÜNVAN"] || '';
        }

        if (name && typeof name === 'string' && name.length > 2) {
          const cleanName = name.replace(/\s*\(.*?\)\s*/g, '').trim();
          dataMap[normalize(cleanName)] = {
            fullName: cleanName,
            photoUrl: getDriveThumbnail(typeof photo === 'string' ? photo : ''),
            role: typeof role === 'string' ? role : '',
            title: typeof title === 'string' ? title : ''
          };
        }
      };

      const processData = (input: any, type: 'pilot' | 'tech') => {
        if (!input) return;
        const arr = Array.isArray(input) ? input : (input.data && Array.isArray(input.data) ? input.data : []);
        arr.forEach(item => processItem(item, type));
      };

      processData(pilotRes, 'pilot');
      processData(techRes, 'tech');
      setPersonnelData(dataMap);
    } catch (err) {
      console.warn('Personnel photo data could not be fetched:', err);
    }
  };

  const fetchExternalData = async () => {
    setIsLoading(true);
    setSyncError(null);
    try {
      // User provided script URL
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

        let lastId = '';
        let lastTarih = '';
        let lastBolge = '';
        
        data.forEach((l, rowIdx) => {
          if (!l) return;
          if (isAOA && rowIdx === 0 && (l[0]?.toString().includes('Sıra') || l[1]?.toString().includes('Tarih'))) return;

          const val = (keys: string[], colIdx?: number) => {
            if (isAOA && Array.isArray(l) && colIdx !== undefined) {
              if (l[colIdx] !== undefined && l[colIdx] !== null && l[colIdx] !== '') return l[colIdx];
            }
            for (const k of keys) {
              if (l[k] !== undefined && l[k] !== null && l[k] !== '') return l[k];
              const target = k.toLowerCase().replace(/[^\w]/g, '');
              for (const actualKey in l) {
                if (actualKey.toLowerCase().replace(/[^\w]/g, '') === target) {
                   if (l[actualKey] !== undefined && l[actualKey] !== null && l[actualKey] !== '') return l[actualKey];
                }
              }
            }
            return null;
          };

          let currentId = (val(['Sıra No', 'id', 'siraNo', 'Log ID', 'No', 'Sira', 'SIRA NO'], 0) || '').toString().trim();
          if (currentId && !currentId.toLowerCase().includes('tarih')) {
            lastId = currentId;
          } else if (!currentId) {
            // If ID is missing, but row has data, generate a temporary row-based ID to ensure it is not skipped
            const rowDataCheck = [2, 3, 4, 10, 17].some(idx => val([], idx));
            if (rowDataCheck) currentId = `row-${rowIdx}`;
          }
          
          const id = currentId || lastId;
          if (!id) return;

          const tarihRaw = val(['Tarih', 'date', 'Log Tarihi', 'TARIH'], 1);
          let currentTarih = '';
          if (tarihRaw) {
            const tStr = tarihRaw.toString().trim();
            if (/^\d{5}$/.test(tStr)) {
              const excelDate = new Date((Number(tStr) - 25569) * 86400 * 1000);
              currentTarih = excelDate.toISOString().split('T')[0];
            } else if (tStr.includes('T')) {
              currentTarih = tStr.split('T')[0];
            } else {
              const parts = tStr.split(/[:./-]/);
              if (parts.length === 3) {
                let p0 = parts[0].padStart(2, '0');
                let p1 = parts[1].padStart(2, '0');
                let p2 = parts[2].trim();
                if (p2.length > 4) p2 = p2.slice(-4);
                if (p2.length === 2) p2 = p2.startsWith('2') ? `20${p2}` : `19${p2}`;
                currentTarih = `${p2}-${p1}-${p0}`;
              } else {
                currentTarih = tStr;
              }
            }
          }
          
          if (currentTarih && currentTarih.includes('-')) {
             const [y, m, d] = currentTarih.split('-');
             if (y && m && d) lastTarih = `${y.padStart(4, '20')}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
          }
          
          const bolgeRaw = val(['Görev Bölgesi', 'gorevBolgesi', 'Region'], 9);
          if (bolgeRaw) lastBolge = bolgeRaw.toString();

          const gorevTipi = val(['Görev Tipi', 'gorevTipi', 'Mission Type'], 8)?.toString() || '';
          
          const cleaned: FlightLog = {
            id,
            tarih: currentTarih || lastTarih || '',
            kaptanPilot: val(['Kaptan Pilot', 'Kaptan', 'Captain', 'Pilot 1', 'PİLOT1', 'kaptanPilot'], 2)?.toString() || '',
            ikinciPilot: val(['2. Pilot', 'Co-Pilot', 'Pilot 2', 'PİLOT2', 'ikinciPilot'], 3)?.toString() || '',
            teknisyen1: val(['Teknisyen', 'Teknisyen 1', 'Tech 1', 'Ekip', 'TEKNİSYEN1', 'teknisyen1'], 4)?.toString() || '',
            operator1: val(['Operator', 'Operatör 1', 'Op 1', 'Operatör', 'OPERATÖR1', 'operator1'], 5)?.toString() || '',
            teknisyen2: val(['Teknisyen', 'Teknisyen 2', 'Teknisyen (2)', 'Tech 2', 'TEKNİSYEN2', 'teknisyen2'], 6)?.toString() || '',
            operator2: val(['Operator', 'Operatör 2', 'Operator (2)', 'Op 2', 'Operatör (2)', 'OPERATÖR2', 'operator2'], 7)?.toString() || '',
            gorevTipi: gorevTipi,
            gorevBolgesi: val(['Görev Bölgesi', 'gorevBolgesi', 'Region'], 9)?.toString() || lastBolge,
            kalkis: formatTimeValue(val(['Kalkış', 'Takeoff', 'kalkis'], 10)),
            inis: formatTimeValue(val(['İniş', 'Landing', 'inis'], 11)),
            ucusSuresi: formatTimeValue(val(['Uçuş Süresi', 'Süre', 'Duration', 'ucusSuresi'], 12)),
            k9YanginHektar: parseNumeric(val(['TK-9 Yangın(Hektar)', 'YANGIN(Hk)', 'Yangın', 'Fire', 'k9YanginHektar'], 13)),
            miktarCekim: parseNumeric(val(['Miktar (Çekim)', 'ÇEKİM', 'Miktar', 'Amount', 'miktarCekim'], 14)),
            tk9GorevHektar: parseNumeric(val(['TK-9 gorev(Hektar)', 'GÖREV(Hk)', 'tk9GorevHektar', 'TK9 Görev'], 15)),
            uyduDk: parseNumeric(val(['Uydu (Dk)', 'UYDU(Dk)', 'uyduDk', 'Uydu dk'], 16)),
            aciklama: val(['AÇIKLAMA', 'AÇIKLAMA (R)', 'Açıklama', 'Remarks', 'Description', 'aciklama'], 17)?.toString() || ''
          };
          
          processedList.push(cleaned);
        });

        const list = append ? [...prev, ...processedList] : processedList;
        
        // Remove duplicates based on all critical fields to be safe
        const uniqueMap = new Map<string, FlightLog>();
        list.forEach(item => {
          const key = `${item.tarih}_${item.id}_${item.kalkis}_${item.kaptanPilot}`;
          uniqueMap.set(key, item);
        });

        return Array.from(uniqueMap.values()).sort((a, b) => {
          const idA = parseInt(a.id) || 0;
          const idB = parseInt(b.id) || 0;
          if (idA !== idB) return idB - idA; // Show newest first in results
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
        log.kaptanPilot, log.ikinciPilot, 
        log.teknisyen1, log.operator1, 
        log.teknisyen2, log.operator2,
        log.aciklama
      ].some(name => isPersonMatch(name || '', filters.pilot));

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
    aoaData.push(['TOPLAM SÜRE', '', '', '', '', '', '', '', '', '', '', '', totalHoursString]);
    const lastDataRowIndex = aoaData.length - 1;
    XLSX.utils.sheet_add_aoa(worksheet, [['TOPLAM SÜRE', '', '', '', '', '', '', '', '', '', '', '', totalHoursString]], { origin: `A${lastDataRowIndex + 1}` });

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

  const exportPersonnelToPDF = () => {
    const doc = new jsPDF();
    const timestamp = new Date().toLocaleDateString('tr-TR');
    
    // Add Turkish fonts support if needed (jspdf supports standard fonts by default)
    // We'll use doc.text with helvetica
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(18);
    doc.setTextColor(16, 185, 129); // emerald-500
    doc.text('OGM B-360 PERSONEL UCUS RAPORU', 105, 15, { align: 'center' });
    
    doc.setFontSize(10);
    doc.setTextColor(100, 116, 139); // slate-500
    doc.text(`Rapor Tarihi: ${timestamp}`, 105, 22, { align: 'center' });

    let currentY = 35;

    const allPersonnel = [...pilotsInResults, ...technicalInResults].filter(p => p && normalize(p).includes(normalize(personnelSearch)));

    if (allPersonnel.length === 0) {
      doc.text('Arama sonuclarina uygun personel bulunamadi.', 105, 50, { align: 'center' });
    } else {
      allPersonnel.forEach((personName, index) => {
        // Check if we need a new page
        if (currentY > 250) {
          doc.addPage();
          currentY = 20;
        }

        const totalHours = calculateTotalHours(logs, personName);
        const stats = getPersonStats(personName, logs, personnelData, pilotsInResults.includes(personName) ? 'Pilot' : 'Teknisyen');

        // Person info block
        doc.setFillColor(248, 250, 252); // slate-50
        doc.rect(14, currentY, 182, 45, 'F');
        doc.setDrawColor(226, 232, 240); // slate-200
        doc.rect(14, currentY, 182, 45, 'S');

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(14);
        doc.setTextColor(30, 41, 59); // slate-800
        doc.text(personName.toUpperCase(), 20, currentY + 10);

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.setTextColor(71, 85, 105); // slate-600
        doc.text(`${stats.role} ${stats.title ? `| ${stats.title}` : ''}`, 20, currentY + 16);

        doc.setFont('helvetica', 'bold');
        doc.text(`TOPLAM UCUS: ${stats.count}`, 150, currentY + 10);
        doc.setTextColor(249, 115, 22); // orange-500
        doc.text(`TOPLAM SURE: ${totalHours}`, 150, currentY + 16);

        // Missions table
        if (stats.distribution.length > 0) {
          const missionData = stats.distribution.map(([type, count]) => [type, count]);
          (doc as any).autoTable({
            startY: currentY + 22,
            head: [['Gorev Tipi', 'Adet']],
            body: missionData,
            margin: { left: 20 },
            tableWidth: 80,
            styles: { fontSize: 8 },
            headStyles: { fillColor: [16, 185, 129] }
          });
          
          // Total Stats Box
          const finalY = (doc as any).lastAutoTable.finalY + 5;
          doc.setFontSize(8);
          doc.setTextColor(30, 41, 59);
          doc.text(`Top. Hektar: ${stats.stats.totalHektar}`, 110, currentY + 28);
          doc.text(`Yangin Hektar: ${stats.stats.yanginHektar}`, 110, currentY + 33);
          doc.text(`Gorev Hektar: ${stats.stats.gorevHektar}`, 110, currentY + 38);
          doc.text(`Top. Uydu (Dk): ${stats.stats.uyduDk}`, 150, currentY + 28);
          doc.text(`Top. Cekim: ${stats.stats.cekim}`, 150, currentY + 33);
        } else {
          doc.setFontSize(8);
          doc.text('Ucus verisi bulunamadi.', 20, currentY + 25);
        }

        currentY += 55;
      });
    }

    doc.save(`Personnel_Report_${personnelSearch || 'Tumu'}.pdf`);
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
          <div className="w-10 h-10 bg-emerald-500 rounded flex items-center justify-center text-forest-base font-black shadow-lg">
            B-360
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
                  onClick={() => setActiveTab('search')}
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
                        setSearchTerm('');
                        setFilters({
                            pilot: '',
                            bolge: '',
                            tip: '',
                            startDate: '',
                            endDate: ''
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
                <div className="bg-forest-dark px-4 py-2 rounded-lg border border-emerald-800 flex flex-col min-w-[120px] text-left">
                  <span className="text-[9px] text-emerald-500/50 uppercase font-black tracking-widest text-left">UYDU / ÇEKİM</span>
                  <div className="flex flex-col">
                    <span className="text-sm font-bold text-blue-400 text-left">
                      {filteredLogs.reduce((acc, l) => acc + (Number(l.uyduDk) || 0), 0).toLocaleString('tr-TR')} <span className="text-[10px]">DK</span>
                    </span>
                  </div>
                </div>
              </div>

              <div className="border border-emerald-800 rounded-lg overflow-hidden flex flex-col bg-forest-dark shadow-2xl">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left min-w-[2800px] table-fixed">
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
                        <td className="px-4 py-4 text-center bg-red-900/20 text-red-500 font-black">
                           {filteredLogs.reduce((acc, l) => acc + (Number(l.k9YanginHektar) || 0), 0).toLocaleString('tr-TR')}
                        </td>
                        <td className="px-4 py-4 text-center text-emerald-300 font-black">
                           {filteredLogs.reduce((acc, l) => acc + (Number(l.miktarCekim) || 0), 0).toLocaleString('tr-TR')}
                        </td>
                        <td className="px-4 py-4 text-center bg-blue-900/20 text-blue-400 font-black">
                           {filteredLogs.reduce((acc, l) => acc + (Number(l.tk9GorevHektar) || 0), 0).toLocaleString('tr-TR')}
                        </td>
                        <td className="px-4 py-4 text-center bg-forest-dark/50"></td>
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
              <div className="bg-forest-dark p-6 rounded-2xl border border-emerald-800 shadow-xl flex flex-col md:flex-row justify-between items-center gap-4">
                <div className="w-full md:w-96">
                   <label className="text-[10px] text-emerald-400 uppercase font-black mb-1.5 block tracking-widest pl-1">Soyisim / İsim ile Sorgula</label>
                   <div className="relative">
                     <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-emerald-700" size={16} />
                     <input 
                       type="text" 
                       placeholder="Aramak istediğiniz personelin soyismini yazın..."
                       className="w-full bg-forest-base border border-emerald-700 rounded px-4 py-3 pl-10 text-sm focus:outline-none focus:border-emerald-500 transition-all font-medium"
                       value={personnelSearch}
                       onChange={(e) => setPersonnelSearch(e.target.value)}
                     />
                   </div>
                </div>
                <div className="flex gap-2 w-full md:w-auto">
                  <button 
                    onClick={exportPersonnelToPDF}
                    className="w-full md:w-auto flex items-center justify-center gap-2 px-8 py-3 bg-red-900/40 border border-red-700 rounded-xl text-red-200 font-black text-[10px] hover:bg-red-800 transition-all active:scale-95 shadow-lg"
                  >
                    <FileText size={14} />
                    PDF İNDİR
                  </button>
                  <button 
                    onClick={fetchExternalData}
                    disabled={isLoading}
                    className="w-full md:w-auto flex items-center justify-center gap-2 px-8 py-3 bg-emerald-900/40 border border-emerald-700 rounded-xl text-emerald-400 font-black text-[10px] hover:bg-emerald-800 transition-all active:scale-95 disabled:opacity-50"
                  >
                    <History size={14} className={isLoading ? 'animate-spin' : ''} />
                    EXCEL VERİLERİNİ GÜNCELLE
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

              <div className="flex justify-end mb-4">
                <button 
                  onClick={() => {
                    if (confirm('Tüm veriler temizlenecek ve sistem yeniden başlatılacak. Emin misiniz?')) {
                      localStorage.clear();
                      window.location.reload();
                    }
                  }}
                  className="bg-red-950/40 hover:bg-red-900/60 text-red-400 border border-red-900/50 px-4 py-2 rounded-lg text-[10px] font-black flex items-center gap-2 transition-all"
                >
                  <Trash2 size={14} /> SİSTEM VERİLERİNİ SIFIRLA
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                 <PersonnelPanel 
                    title="Uçuş Pilotları" 
                    data={PILOTS.filter(p => normalize(p).includes(normalize(personnelSearch)))} 
                    type="Pilot" 
                    logs={logs} 
                    personnelData={personnelData}
                 />
                 <PersonnelPanel 
                    title="Teknik Ekip & Operatörler" 
                    data={TECHNICIANS.filter(p => normalize(p).includes(normalize(personnelSearch)))} 
                    type="Teknisyen" 
                    logs={logs} 
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
          <span className="flex items-center gap-2"><Clock size={14} className="opacity-50" /> SON KAYIT: {formatDisplayDate(logs[logs.length - 1]?.tarih) || 'N/A'}</span>
        </div>
        <p className="text-[10px] opacity-40 font-mono tracking-widest">© 2024 Havacılık Bilgi Sistemleri | Secure B-360 DB</p>
      </footer>
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
                    <h3 className="font-bold text-sm tracking-wide uppercase italic leading-tight text-white">{name}</h3>
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
      if (!formData.teknisyen1 || !formData.operator1) {
        alert('Teknik ekipte Teknisyen 1 ve Operatör 1 zorunludur.');
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

