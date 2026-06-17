export interface FlightLog {
  id: string;
  tarih: string;
  kaptanPilot: string;
  ikinciPilot: string;
  teknisyen1: string;
  operator1: string;
  teknisyen2: string;
  operator2: string;
  gorevTipi: string;
  gorevBolgesi: string;
  kalkis: string; // HH:mm
  inis: string;  // HH:mm
  ucusSuresi: string; // Calculated
  k9YanginHektar: number;
  miktarCekim: number;
  tk9GorevHektar: number;
  uyduDk: number | string;
  aciklama: string;
}

export interface Personnel {
  id: string;
  name: string;
  role: 'Pilot' | 'Teknisyen';
  fullName: string;
}

export const PILOTS = [
  'YILMAZ MAMUNLUOĞLU',
  'AYDIN TÜTÜNCÜOĞLU',
  'ALTAN ALKAN SÖZEN',
  'CENGİZ ÖZDEMİR',
  'DEVRİM FERHAT ÇALIŞKAN',
  'SERKAN KEBABCI',
  'AYDEMİR TEZGEL'
];

export const TECHNICIANS = [
  'TEZCAN GÜZER',
  'SERKAN KEBAPCI',
  'HASAN AKSOY',
  'FERHAT ÖZCAN',
  'ÖMER ERSOY',
  'AYCAN TAN'
];

export const GOREV_TIPLERI = [
  'TK-9 Çekimi',
  'Yangın Uçuşu',
  'PMN(Personel Malzeme Nakli)',
  'VIP',
  'Eğitim',
  'İntibak Kursu',
  'TK-9 Eğitimi',
  'Tatbikat',
  'Yurt Dışı Yangın',
  'Diğer Kurumlar TK-9',
  'Kamera Testi',
  'İntikal Uçuşu',
  'Dogal Afet',
  'Yer Calistirmasi'
];
