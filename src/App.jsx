import { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { jsPDF } from 'jspdf';
import { autoCropDocument } from './WarpHelper.js';
import { MapPin, Flag, Save, Download, Trash2, Loader2, AlertCircle, X, Navigation, ChevronDown, ChevronUp, ArrowDown, DollarSign, Calendar, TrendingUp, History, Sun, Moon, Volume2, VolumeX, Vibrate, Camera, Settings, Receipt, ParkingCircle, Eye, Check, Image as ImageIcon, FileText, Aperture, Pencil } from 'lucide-react';
import * as XLSX from 'xlsx';

// ═══════════════════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════════════════
const RATE = 1.14; // R$ / km
const APP_VERSION = 'v3·1';

const SEED_TRIPS = [
  { id: 's1', date: '20/04/2026', origin: 'Rua Attilio Ceccarelli, 90 - Jardim Rio Pequeno, São Paulo - SP, 05388-040', destination: 'Rua dos Marianos, 349 - Centro, Osasco - SP, 06016-050', km: null, geometry: null },
  { id: 's2', date: '20/04/2026', origin: 'Rua dos Marianos, 349 - Centro, Osasco - SP, 06016-050', destination: 'Avenida Marechal Rondon, 199 - Centro, Osasco - SP, 06093-020', km: null, geometry: null },
  { id: 's3', date: '20/04/2026', origin: 'Avenida Marechal Rondon, 165 - Centro, Osasco - SP, 06093-020', destination: 'Rua Antônio Agú, 833 - Centro, Osasco - SP, 06013-000', km: null, geometry: null },
  { id: 's4', date: '20/04/2026', origin: 'Rua Antônio Agú, 833 - Centro, Osasco - SP, 06013-000', destination: 'Rua Minas Bogasian, 284 - Centro, Osasco - SP, 06013-010', km: null, geometry: null },
  { id: 's5', date: '20/04/2026', origin: 'Rua Minas Bogasian, 284 - Centro, Osasco - SP, 06013-010', destination: 'Rua Dona Primitiva Vianco, 589 - Centro, Osasco - SP, 06010-004', km: null, geometry: null },
  { id: 's6', date: '20/04/2026', origin: 'Rua Dona Primitiva Vianco, 589 - Centro, Osasco - SP, 06010-004', destination: 'Avenida João Batista, 11 - Centro, Osasco - SP, 06097-100', km: null, geometry: null },
  { id: 's7', date: '20/04/2026', origin: 'Avenida João Batista, 11 - Centro, Osasco - SP, 06097-100', destination: 'Avenida João Batista, 253 - Centro, Osasco - SP, 06090-100', km: null, geometry: null },
  { id: 's8', date: '20/04/2026', origin: 'Avenida João Batista, 253 - Centro, Osasco - SP, 06090-100', destination: 'Rua Fiorino Beltrano, 195 - Centro, Osasco - SP, 06097-040', km: null, geometry: null },
  { id: 's9', date: '20/04/2026', origin: 'Rua Fiorino Beltrano, 195 - Centro, Osasco - SP, 06097-040', destination: 'Rua Attilio Ceccarelli, 90 - Jardim Rio Pequeno, São Paulo - SP, 05388-040', km: null, geometry: null },
];

const STATE_MAP = {'Acre':'AC','Alagoas':'AL','Amapá':'AP','Amazonas':'AM','Bahia':'BA','Ceará':'CE','Distrito Federal':'DF','Espírito Santo':'ES','Goiás':'GO','Maranhão':'MA','Mato Grosso':'MT','Mato Grosso do Sul':'MS','Minas Gerais':'MG','Pará':'PA','Paraíba':'PB','Paraná':'PR','Pernambuco':'PE','Piauí':'PI','Rio de Janeiro':'RJ','Rio Grande do Norte':'RN','Rio Grande do Sul':'RS','Rondônia':'RO','Roraima':'RR','Santa Catarina':'SC','São Paulo':'SP','Sergipe':'SE','Tocantins':'TO'};

const MONTH_NAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const MONTH_FULL = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

// ═══════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function distMeters(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
function haversineKm(lat1, lon1, lat2, lon2) { return distMeters(lat1,lon1,lat2,lon2) / 1000; }

async function geocodeAddress(address) {
  const q = encodeURIComponent(address);
  const url = `https://nominatim.openstreetmap.org/search?q=${q}&format=json&accept-language=pt-BR&limit=1&countrycodes=br`;
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const list = await r.json();
    if (list.length > 0) {
      return { lat: parseFloat(list[0].lat), lng: parseFloat(list[0].lon) };
    }
    return null;
  } catch { return null; }
}

function formatBrazilianAddress(data, overrideNumber, isApprox) {
  const a = data.address || {};
  const street = a.road || a.pedestrian || a.path || '';
  const number = overrideNumber || a.house_number || '';
  const neighborhood = a.suburb || a.neighbourhood || a.quarter || a.city_district || '';
  const city = a.city || a.town || a.village || a.municipality || '';
  const stateAbbr = STATE_MAP[a.state || ''] || a.state || '';
  const postcode = a.postcode || '';
  let parts = [];
  if (street) {
    let nl = number ? (isApprox ? `~${number}` : `${number}`) : '';
    let s = nl ? `${street}, ${nl}` : street;
    if (neighborhood) s += ` - ${neighborhood}`;
    parts.push(s);
  } else if (neighborhood) parts.push(neighborhood);
  let loc = city;
  if (stateAbbr) loc += (loc ? ' - ' : '') + stateAbbr;
  if (postcode) loc += (loc ? ', ' : '') + postcode;
  if (loc) parts.push(loc);
  return parts.length > 0 ? parts.join(', ') : (data.display_name || '');
}

async function reverseLookup(lat, lng, zoom = 19) {
  const url = `https://nominatim.openstreetmap.org/reverse?lat=${lat}&lon=${lng}&format=json&accept-language=pt-BR&addressdetails=1&zoom=${zoom}`;
  const r = await fetch(url);
  if (!r.ok) throw new Error('Erro no serviço de endereços');
  return await r.json();
}

async function searchNearbyForNumber(lat, lng, road) {
  if (!road) return null;
  const delta = 0.0015;
  const viewbox = `${lng-delta},${lat-delta},${lng+delta},${lat+delta}`;
  const q = encodeURIComponent(road);
  const url = `https://nominatim.openstreetmap.org/search?q=${q}&format=json&addressdetails=1&accept-language=pt-BR&limit=20&bounded=1&viewbox=${viewbox}`;
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const list = await r.json();
    if (!Array.isArray(list) || list.length === 0) return null;
    const candidates = list
      .filter(item => item.address?.house_number && (item.address?.road === road || item.address?.pedestrian === road))
      .map(item => ({ number: item.address.house_number, d: distMeters(lat, lng, parseFloat(item.lat), parseFloat(item.lon)) }))
      .sort((a, b) => a.d - b.d);
    return candidates[0]?.number || null;
  } catch { return null; }
}

async function getAddressWithNumber(lat, lng) {
  const data = await reverseLookup(lat, lng, 19);
  const direct = data?.address?.house_number;
  if (direct) return { data, number: direct, isApprox: false };
  const road = data?.address?.road || data?.address?.pedestrian;
  const nearby = await searchNearbyForNumber(lat, lng, road);
  if (nearby) return { data, number: nearby, isApprox: true };
  return { data, number: null, isApprox: false };
}

// BRL formatter with parts (for typographic split)
const brlFmt = new Intl.NumberFormat('pt-BR', { style:'currency', currency:'BRL', minimumFractionDigits:2, maximumFractionDigits:2 });
const numFmt = new Intl.NumberFormat('pt-BR', { minimumFractionDigits:2, maximumFractionDigits:2 });

// localStorage
const TRIPS_KEY = 'km_trips_v1';
const THEME_KEY = 'coca_theme';
const SOUND_KEY = 'coca_sound';
const HAPTIC_KEY = 'coca_haptic';
const GEMINI_KEY = 'gemini_api_key';
const RECEIPTS_META_KEY = 'km_receipts_meta_v1';
const loadTrips = () => { try { const r = localStorage.getItem(TRIPS_KEY); return r ? JSON.parse(r) : null; } catch { return null; } };
const saveTrips = (t) => { try { localStorage.setItem(TRIPS_KEY, JSON.stringify(t)); } catch {} };
const loadReceiptsMeta = () => { try { const r = localStorage.getItem(RECEIPTS_META_KEY); return r ? JSON.parse(r) : []; } catch { return []; } };
const saveReceiptsMeta = (m) => { try { localStorage.setItem(RECEIPTS_META_KEY, JSON.stringify(m)); } catch {} };

// ═══════════════════════════════════════════════════════════════════════════
// INDEXED DB (for receipt images — localStorage can't hold large blobs)
// ═══════════════════════════════════════════════════════════════════════════
const DB_NAME = 'km_tracker_db';
const DB_VERSION = 1;
const STORE_NAME = 'receipt_images';

function openReceiptDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveReceiptImage(id, blob) {
  const db = await openReceiptDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put({ id, blob });
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function getReceiptImage(id) {
  const db = await openReceiptDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).get(id);
    req.onsuccess = () => resolve(req.result?.blob || null);
    req.onerror = () => reject(req.error);
  });
}

async function deleteReceiptImage(id) {
  const db = await openReceiptDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// GEMINI VISION API
// ═══════════════════════════════════════════════════════════════════════════
async function getBestGeminiModel(apiKey) {
  try {
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
    if (!response.ok) return 'models/gemini-1.5-flash';
    
    const data = await response.json();
    if (!data.models) return 'models/gemini-1.5-flash';

    const availableNames = data.models.map(m => m.name);
    // Removemos os modelos 2.0 pois a API lista eles, mas bloqueia o uso para contas novas/gratuitas
    const preferences = [
      "models/gemini-1.5-flash",
      "models/gemini-1.5-pro",
      "models/gemini-1.5-flash-8b"
    ];

    for (const pref of preferences) {
      if (availableNames.includes(pref)) return pref;
    }
    
    const fallback = data.models.find(m => m.name.includes("gemini") && 
      (m.supportedGenerationMethods?.includes("generateContent") || m.supportedMethods?.includes("generateContent"))
    );
    if (fallback) return fallback.name;
    
    return 'models/gemini-1.5-flash';
  } catch (e) {
    return 'models/gemini-1.5-flash';
  }
}

async function analyzeReceiptWithGemini(imageBase64, apiKey) {
  const modelName = await getBestGeminiModel(apiKey);
  const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${apiKey}`;
  const prompt = `Analise esta foto de um comprovante de pedágio ou estacionamento brasileiro.
Extraia os seguintes dados em formato JSON puro (sem markdown, sem \`\`\`json):
{
  "tipo": "pedagio" ou "estacionamento",
  "concessionaria": "nome da concessionária ou estabelecimento",
  "cnpj": "CNPJ se visível ou null",
  "dataHora": "DD/MM/AAAA HH:MM:SS",
  "praca": "praça do pedágio ou local do estacionamento",
  "via": "rodovia ou endereço",
  "placa": "placa do veículo se visível ou null",
  "classe": "classe do veículo se visível ou null",
  "valor": 0.00,
  "recibo": "número do recibo se visível ou null",
  "dfe": "string do DFE (Documento Fiscal Equivalente) exatamente como impresso, garantindo o formato de texto para manter os zeros à esquerda, ou null"
}
Se algum campo não estiver visível no documento, retorne null para esse campo.
Retorne SOMENTE o JSON puro, sem explicações, sem formatação markdown.`;

  const body = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'image/jpeg', data: imageBase64 } }
      ]
    }]
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro na API Gemini (${response.status})`);
  }

  const data = await response.json();
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  // Clean markdown code fences if present
  const cleaned = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    throw new Error('A IA não conseguiu extrair os dados. Tente tirar outra foto com melhor iluminação.');
  }
}

/**
 * Inteligência 2.0: Resolve nomes de empresas + cidades/bairros
 * Agora retorna uma LISTA de sugestões para o dropdown.
 */
async function resolvePlaceWithGemini(query, apiKey) {
  if (!query || query.length < 3) return [];
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const prompt = `Você é um resolvedor de endereços geográficos especializado no Brasil.
Seu objetivo é fornecer uma lista de sugestões de endereços reais baseados em buscas por "Nome de Empresa + Localidade" ou endereços incompletos.

ENTRADA: "${query}"

REGRAS:
1. Retorne até 5 sugestões que façam sentido no contexto geográfico (Cidade, Bairro) mencionado.
2. Cada sugestão deve incluir o endereço formatado (com número se possível), latitude e longitude.
3. Se o usuário digitar "FEMSA Ponta Grossa", a lista deve conter as unidades da FEMSA nessa cidade.
4. Responda APENAS um JSON no formato: [{"address": "Rua..., Num - Bairro, Cidade - UF, CEP", "lat": -23.123, "lng": -46.123}, ...]
5. Não use markdown, responda apenas o array JSON.`;

  const body = { contents: [{ parts: [{ text: prompt }] }] };
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!response.ok) return [];
    const data = await response.json();
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
    const cleaned = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const parsed = JSON.parse(cleaned);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// FEEDBACK (sound + haptic)
// ═══════════════════════════════════════════════════════════════════════════
let audioCtx;
function getAudioCtx() {
  if (typeof window === 'undefined') return null;
  if (!audioCtx) {
    try { audioCtx = new (window.AudioContext || window.webkitAudioContext)(); } catch { return null; }
  }
  return audioCtx;
}

function playTone(freq, duration = 0.04, type = 'sine', volume = 0.15) {
  const ctx = getAudioCtx();
  if (!ctx) return;
  const osc = ctx.createOscillator();
  const gain = ctx.createGain();
  osc.type = type;
  osc.frequency.value = freq;
  gain.gain.setValueAtTime(0, ctx.currentTime);
  gain.gain.linearRampToValueAtTime(volume, ctx.currentTime + 0.005);
  gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + duration);
  osc.connect(gain);
  gain.connect(ctx.destination);
  osc.start();
  osc.stop(ctx.currentTime + duration);
}

const SOUND_PRESETS = {
  tap:        () => playTone(1000, 0.025, 'sine', 0.10),
  success:    () => { playTone(660, 0.06, 'sine', 0.12); setTimeout(()=>playTone(880, 0.10, 'sine', 0.14), 60); },
  error:      () => { playTone(400, 0.08, 'sawtooth', 0.10); setTimeout(()=>playTone(280, 0.12, 'sawtooth', 0.10), 80); },
  swoosh:     () => playTone(800, 0.05, 'triangle', 0.08),
};
const HAPTIC_PRESETS = { tap: 10, success: [15, 50, 15], error: [50, 30, 50, 30, 50], swoosh: 8 };

function useFeedback() {
  return useCallback((key) => {
    if (typeof window === 'undefined') return;
    if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    if (localStorage.getItem(SOUND_KEY) !== 'off') {
      try { SOUND_PRESETS[key]?.(); } catch {}
    }
    if (localStorage.getItem(HAPTIC_KEY) !== 'off' && 'vibrate' in navigator) {
      try { navigator.vibrate(HAPTIC_PRESETS[key] || 10); } catch {}
    }
  }, []);
}

// ═══════════════════════════════════════════════════════════════════════════
// THEME MANAGEMENT (anti-FOUC handled in index.html via inline script)
// ═══════════════════════════════════════════════════════════════════════════
function getInitialTheme() {
  if (typeof window === 'undefined') return 'light';
  const stored = localStorage.getItem(THEME_KEY);
  if (stored === 'dark' || stored === 'light') return stored;
  return matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function setTheme(next) {
  document.documentElement.setAttribute('data-theme', next);
  document.documentElement.style.colorScheme = next;
  localStorage.setItem(THEME_KEY, next);
  // Update theme-color meta
  const metaLight = document.querySelector('meta[name="theme-color"][media*="light"]');
  const metaDark  = document.querySelector('meta[name="theme-color"][media*="dark"]');
  if (metaLight) metaLight.setAttribute('content', next === 'dark' ? '#0A0A0A' : '#FAFAFA');
  if (metaDark)  metaDark.setAttribute('content', next === 'dark' ? '#0A0A0A' : '#FAFAFA');
}

function toggleTheme(current) {
  const next = current === 'dark' ? 'light' : 'dark';
  if (typeof document !== 'undefined' && document.startViewTransition) {
    document.startViewTransition(() => setTheme(next));
  } else {
    setTheme(next);
  }
  return next;
}

// ═══════════════════════════════════════════════════════════════════════════
// MINI MAP (brutalist SVG)
// ═══════════════════════════════════════════════════════════════════════════
function MiniMap({ geometry }) {
  if (!geometry || geometry.length < 2) {
    return (
      <div className="km-mini-empty">
        <span>◊ rota indisponível ◊</span>
      </div>
    );
  }
  let minLng=Infinity, maxLng=-Infinity, minLat=Infinity, maxLat=-Infinity;
  for (const [lng, lat] of geometry) {
    if (lng < minLng) minLng = lng; if (lng > maxLng) maxLng = lng;
    if (lat < minLat) minLat = lat; if (lat > maxLat) maxLat = lat;
  }
  const padX = ((maxLng - minLng) || 0.001) * 0.15;
  const padY = ((maxLat - minLat) || 0.001) * 0.20;
  minLng -= padX; maxLng += padX; minLat -= padY; maxLat += padY;
  const W=320, H=100;
  const dataAspect = (maxLng - minLng) / (maxLat - minLat);
  const viewAspect = W / H;
  if (dataAspect > viewAspect) {
    const r = (maxLng - minLng) / viewAspect;
    const m = (minLat + maxLat) / 2; minLat = m - r/2; maxLat = m + r/2;
  } else {
    const r = (maxLat - minLat) * viewAspect;
    const m = (minLng + maxLng) / 2; minLng = m - r/2; maxLng = m + r/2;
  }
  const project = ([lng, lat]) => [((lng - minLng) / (maxLng - minLng)) * W, H - ((lat - minLat) / (maxLat - minLat)) * H];
  const points = geometry.map(project);
  const pathD = points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${x.toFixed(1)},${y.toFixed(1)}`).join(' ');
  const [sx, sy] = points[0];
  const [ex, ey] = points[points.length - 1];
  return (
    <div className="km-minimap">
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="xMidYMid slice">
        <defs>
          <pattern id="mapgrid" width="16" height="16" patternUnits="userSpaceOnUse">
            <path d="M 16 0 L 0 0 0 16" fill="none" stroke="var(--map-grid)" strokeWidth="0.5" />
          </pattern>
          <linearGradient id="routegrad" x1="0%" y1="0%" x2="100%" y2="0%">
            <stop offset="0%" stopColor="var(--coca-red)" />
            <stop offset="100%" stopColor="var(--text-primary)" />
          </linearGradient>
        </defs>
        <rect width={W} height={H} fill="url(#mapgrid)" />
        <path d={pathD} stroke="var(--coca-red-glow)" strokeWidth="6" fill="none" strokeLinecap="round" strokeLinejoin="round" />
        <path d={pathD} stroke="url(#routegrad)" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" className="km-route-dash" />
        <circle cx={sx} cy={sy} r="5" fill="var(--bg-base)" stroke="var(--coca-red)" strokeWidth="2" />
        <circle cx={sx} cy={sy} r="2" fill="var(--coca-red)" />
        <rect x={ex - 4} y={ey - 4} width="8" height="8" fill="var(--text-primary)" stroke="var(--bg-base)" strokeWidth="1.5" />
      </svg>
      <div className="km-minimap-label">
        <span>◊ ROTA</span>
        <span>{geometry.length} pts</span>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// ACTIVITY RINGS (Apple-style)
// ═══════════════════════════════════════════════════════════════════════════
function ActivityRings({ kmProgress, moneyProgress, daysProgress }) {
  const rings = [
    { progress: kmProgress, color: 'var(--coca-red)', radius: 70, label: 'KM' },
    { progress: moneyProgress, color: 'var(--text-primary)', radius: 54, label: 'R$' },
    { progress: daysProgress, color: 'var(--text-secondary)', radius: 38, label: 'DIAS' },
  ];
  const stroke = 11;
  return (
    <div className="km-rings">
      <svg viewBox="0 0 200 200" width="180" height="180">
        {rings.map((r, i) => {
          const c = 2 * Math.PI * r.radius;
          return (
            <g key={i}>
              <circle r={r.radius} cx="100" cy="100" fill="none" stroke="var(--border-subtle)" strokeWidth={stroke} />
              <circle
                r={r.radius} cx="100" cy="100" fill="none"
                stroke={r.color} strokeWidth={stroke}
                strokeDasharray={c}
                strokeDashoffset={c * (1 - Math.min(r.progress, 1))}
                strokeLinecap="round"
                transform="rotate(-90 100 100)"
                style={{ transition: 'stroke-dashoffset 1.2s cubic-bezier(.32,.72,0,1)' }}
              />
            </g>
          );
        })}
      </svg>
      <div className="km-rings-legend">
        {rings.map((r, i) => (
          <div key={i} className="km-rings-legend-row">
            <span className="km-rings-dot" style={{ background: r.color }} />
            <span className="km-rings-legend-label">{r.label}</span>
            <span className="km-rings-legend-pct">{Math.round(r.progress * 100)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SPARKLINE
// ═══════════════════════════════════════════════════════════════════════════
function Sparkline({ values, height = 36, labels }) {
  const [touchIdx, setTouchIdx] = useState(null);
  const svgRef = useRef(null);
  if (!values || values.length === 0) return null;
  const W = 280, H = height;
  const max = Math.max(...values, 0.01);
  const step = W / Math.max(values.length - 1, 1);
  const points = values.map((v, i) => [i * step, H - (v / max) * (H - 4) - 2]);
  const pathD = points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${x.toFixed(1)},${y.toFixed(1)}`).join(' ');
  const areaD = pathD + ` L${W},${H} L0,${H} Z`;

  function handleTouch(e) {
    if (!svgRef.current) return;
    const rect = svgRef.current.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const relX = (clientX - rect.left) / rect.width;
    const idx = Math.round(relX * (values.length - 1));
    if (idx >= 0 && idx < values.length) setTouchIdx(idx);
  }

  return (
    <div className="km-sparkline-wrap" style={{ position: 'relative' }}>
      <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`} className="km-sparkline" preserveAspectRatio="none"
        onTouchStart={handleTouch} onTouchMove={handleTouch} onTouchEnd={() => setTouchIdx(null)}
        onMouseMove={handleTouch} onMouseLeave={() => setTouchIdx(null)}
      >
        <defs>
          <linearGradient id="sparkfill" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="var(--coca-red)" stopOpacity="0.3" />
            <stop offset="100%" stopColor="var(--coca-red)" stopOpacity="0" />
          </linearGradient>
        </defs>
        <path d={areaD} fill="url(#sparkfill)" />
        <path d={pathD} fill="none" stroke="var(--coca-red)" strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round" />
        {touchIdx != null && points[touchIdx] && (
          <>
            <line x1={points[touchIdx][0]} y1={0} x2={points[touchIdx][0]} y2={H} stroke="var(--coca-red)" strokeWidth="1" strokeDasharray="3,3" opacity="0.6" />
            <circle cx={points[touchIdx][0]} cy={points[touchIdx][1]} r="4" fill="var(--coca-red)" stroke="var(--bg-base)" strokeWidth="2" />
          </>
        )}
      </svg>
      {touchIdx != null && values[touchIdx] != null && (
        <div className="km-sparkline-tooltip" style={{ left: `${(touchIdx / (values.length - 1)) * 100}%` }}>
          <span className="km-mono km-mono-tiny">{numFmt.format(values[touchIdx])} KM</span>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// BIG NUMBER (R$ split — Stripe/Wise style)
// ═══════════════════════════════════════════════════════════════════════════
function BigCurrency({ value, size = 'xl' }) {
  const parts = brlFmt.formatToParts(value);
  let symbol = '', integer = '', decimal = '';
  for (const p of parts) {
    if (p.type === 'currency') symbol = p.value;
    else if (p.type === 'integer' || p.type === 'group') integer += p.value;
    else if (p.type === 'decimal') decimal += p.value;
    else if (p.type === 'fraction') decimal += p.value;
  }
  return (
    <span className={`km-big-currency km-big-currency--${size}`}>
      <span className="km-big-currency-symbol">{symbol}</span>
      <span className="km-big-currency-int">{integer}</span>
      <span className="km-big-currency-dec">{decimal}</span>
    </span>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// AUTOCOMPLETE LIST
// ═══════════════════════════════════════════════════════════════════════════
function SuggestionsList({ suggestions, onSelect, loading }) {
  if (loading && suggestions.length === 0) return null;
  if (!loading && suggestions.length === 0) return null;

  return (
    <div className="km-suggestions">
      {suggestions.map((s, i) => (
        <div key={i} className="km-suggestion-item km-press" onClick={() => onSelect(s)}>
          <div className="km-suggestion-icon">
            <MapPin size={14} />
          </div>
          <div className="km-suggestion-content">
            <div className="km-suggestion-addr">{s.address}</div>
            <div className="km-suggestion-meta">SUGESTÃO · GPS OK</div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function Dashboard({ trips, receipts = [] }) {
  const stats = useMemo(() => {
    const totalKm = trips.reduce((s, t) => s + (t.km || 0), 0);
    const measured = trips.filter(t => t.km != null);

    const pedagios = receipts.filter(r => r.tipo === 'pedagio');
    const estacionamentos = receipts.filter(r => r.tipo === 'estacionamento');
    const totalPedagios = pedagios.reduce((s, r) => s + (r.valor || 0), 0);
    const totalEstacionamentos = estacionamentos.reduce((s, r) => s + (r.valor || 0), 0);
    const totalDespesas = totalPedagios + totalEstacionamentos;

    const byDay = {};
    for (const t of measured) {
      if (!byDay[t.date]) byDay[t.date] = { km: 0, count: 0, toll: 0, parking: 0 };
      byDay[t.date].km += t.km;
      byDay[t.date].count += 1;
    }
    for (const r of receipts) {
      const date = r.dataHora?.split(' ')[0] || '';
      if (!date) continue;
      if (!byDay[date]) byDay[date] = { km: 0, count: 0, toll: 0, parking: 0 };
      if (r.tipo === 'pedagio') byDay[date].toll += (r.valor || 0);
      else byDay[date].parking += (r.valor || 0);
    }

    const days = Object.entries(byDay).map(([date, d]) => ({ date, ...d }))
      .sort((a, b) => {
        const p = s => { const [d,m,y] = s.split('/'); return new Date(y,m-1,d); };
        return p(b.date) - p(a.date);
      });

    const byMonth = {};
    for (const t of measured) {
      const [d, m, y] = t.date.split('/');
      const key = `${m}/${y}`;
      if (!byMonth[key]) byMonth[key] = { km: 0, count: 0, days: new Set(), toll: 0, parking: 0 };
      byMonth[key].km += t.km;
      byMonth[key].count += 1;
      byMonth[key].days.add(t.date);
    }
    for (const r of receipts) {
      const date = r.dataHora?.split(' ')[0] || '';
      if (!date) continue;
      const [d, m, y] = date.split('/');
      const key = `${m}/${y}`;
      if (!byMonth[key]) byMonth[key] = { km: 0, count: 0, days: new Set(), toll: 0, parking: 0 };
      if (r.tipo === 'pedagio') byMonth[key].toll += (r.valor || 0);
      else byMonth[key].parking += (r.valor || 0);
    }

    const months = Object.entries(byMonth).map(([key, d]) => ({ key, ...d, daysCount: d.days.size }))
      .sort((a, b) => {
        const [ma, ya] = a.key.split('/');
        const [mb, yb] = b.key.split('/');
        return new Date(yb, mb-1) - new Date(ya, ma-1);
      });

    const maxDayKm = Math.max(...days.map(d => d.km), 0.01);
    const avgPerDay = days.length > 0 ? totalKm / days.length : 0;
    const bestDay = days.reduce((a, b) => (a?.km || 0) > b.km ? a : b, null);

    const now = new Date();
    const currentKey = `${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`;
    const currentMonth = byMonth[currentKey];
    const daysInMonth = new Date(now.getFullYear(), now.getMonth()+1, 0).getDate();
    const today = now.getDate();
    let projection = null;
    if (currentMonth && today > 0) {
      const projKm = (currentMonth.km / Math.min(today, daysInMonth)) * daysInMonth;
      projection = { km: projKm, money: (projKm * RATE) + currentMonth.toll + currentMonth.parking };
    }

    const last14 = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(now);
      d.setDate(now.getDate() - i);
      const key = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
      last14.push(byDay[key]?.km || 0);
    }

    return { totalKm, totalPedagios, totalEstacionamentos, totalDespesas, measuredCount: measured.length, days, months, maxDayKm, avgPerDay, bestDay, projection, last14, currentMonth };
  }, [trips, receipts]);

  const totalEarning = (stats.totalKm * RATE) + stats.totalDespesas;

  return (
    <div className="km-dash km-fade-up">
      {/* HERO METRIC */}
      <section className="km-card km-card--hero">
        <div className="km-aurora" />
        <header className="km-card-head">
          <span className="km-mono km-mono-label">[ TOTAL GERAL A RECEBER ]</span>
          <span className="km-mono km-mono-label">{String(trips.length).padStart(3,'0')}/REC</span>
        </header>
        <div className="km-hero-amount">
          <BigCurrency value={totalEarning} size="xl" />
        </div>
        <div className="km-hero-meta">
          <span className="km-mono">{numFmt.format(stats.totalKm)} KM (R$ {brlFmt.format(stats.totalKm * RATE)})</span>
          {stats.totalDespesas > 0 && (
            <>
              <span className="km-divider-vert" />
              <span className="km-mono" style={{ color: 'var(--coca-red)' }}>+ DESP: R$ {brlFmt.format(stats.totalDespesas)}</span>
            </>
          )}
        </div>
        {stats.last14.some(v => v > 0) && (
          <div className="km-hero-spark">
            <Sparkline values={stats.last14} height={32} />
            <span className="km-mono km-mono-tiny">14d</span>
          </div>
        )}
      </section>

      {/* PROJECTION + STATS GRID */}
      {stats.projection && (
        <section className="km-card km-card--projection">
          <div className="km-row-tight">
            <span className="km-mono km-mono-label">◊ PROJEÇÃO · {MONTH_FULL[parseInt(stats.projection ? stats.currentMonth ? Object.keys({})[0] : '01' : '01')-1]?.toUpperCase() || ''}{(() => { const n = new Date(); return `${MONTH_FULL[n.getMonth()].toUpperCase()} ${n.getFullYear()}`; })()}</span>
          </div>
          <div className="km-projection-content">
            <div>
              <BigCurrency value={stats.projection.money} size="md" />
              <div className="km-mono km-muted km-mono-tiny km-mt-1">{numFmt.format(stats.projection.km)} KM ESTIMADOS</div>
            </div>
          </div>
        </section>
      )}

      {/* RINGS */}
      <section className="km-card km-card--rings">
        <header className="km-card-head">
          <span className="km-mono km-mono-label">◊ ATIVIDADE · MÊS</span>
        </header>
        <ActivityRings
          kmProgress={Math.min((stats.currentMonth?.km || 0) / 1500, 1)}
          moneyProgress={Math.min((stats.currentMonth?.km || 0) * RATE / 2000, 1)}
          daysProgress={Math.min((stats.currentMonth?.daysCount || 0) / 22, 1)}
        />
      </section>

      {/* STATS GRID */}
      <section className="km-stats-grid">
        <div className="km-stat-tile">
          <span className="km-mono km-mono-label">[ MÉDIA / DIA ]</span>
          <span className="km-stat-num">{numFmt.format(stats.avgPerDay)}</span>
          <span className="km-mono km-muted km-mono-tiny">KM</span>
        </div>
        <div className="km-stat-tile">
          <span className="km-mono km-mono-label">[ MELHOR DIA ]</span>
          <span className="km-stat-num">{stats.bestDay ? numFmt.format(stats.bestDay.km) : '0'}</span>
          <span className="km-mono km-muted km-mono-tiny">{stats.bestDay?.date || '—'}</span>
        </div>
      </section>

      {/* PULL QUOTE EDITORIAL (1× per screen) */}
      {stats.bestDay && stats.totalKm > 0 && (
        <blockquote className="km-pullquote">
          <p>
            <span className="km-serif km-italic">Maior consumo da rota em </span>
            <span className="km-mono km-pullquote-data">{stats.bestDay.date.slice(3)}</span>
            <span className="km-serif km-italic">.</span>
          </p>
          <cite className="km-mono km-mono-tiny km-muted">— REGISTRO INTERNO · DRIVE LOG</cite>
        </blockquote>
      )}

      {/* BY DAY */}
      <section className="km-card">
        <header className="km-card-head km-card-head--bordered">
          <div className="km-row-tight">
            <Calendar size={14} strokeWidth={2.4} />
            <span className="km-mono km-mono-label">POR DIA</span>
          </div>
          <span className="km-mono km-mono-label">{String(stats.days.length).padStart(2,'0')}</span>
        </header>
        {stats.days.length === 0 ? (
          <p className="km-empty">Sem dados ainda.</p>
        ) : (
          <ul className="km-rows">
            {stats.days.map((d) => {
              const earning = d.km * RATE;
              const pct = (d.km / stats.maxDayKm) * 100;
              return (
                <li key={d.date} className="km-row">
                  <div className="km-row-main">
                    <div>
                      <div className="km-mono km-row-date">{d.date}</div>
                      <div className="km-mono km-mono-tiny km-muted">{d.count} {d.count === 1 ? 'VIAGEM' : 'VIAGENS'}</div>
                    </div>
                    <div className="km-row-right">
                      <div className="km-mono km-row-km">{numFmt.format(d.km)} KM</div>
                      <div className="km-row-money">{brlFmt.format(earning + d.toll + d.parking)}</div>
                    </div>
                  </div>
                  <div className="km-bar"><div className="km-bar-fill" style={{ width: `${pct}%` }} /></div>
                </li>
              );
            })}
          </ul>
        )}
      </section>

      {/* BY MONTH */}
      <section className="km-card">
        <header className="km-card-head km-card-head--bordered">
          <div className="km-row-tight">
            <TrendingUp size={14} strokeWidth={2.4} />
            <span className="km-mono km-mono-label">POR MÊS</span>
          </div>
          <span className="km-mono km-mono-label">{String(stats.months.length).padStart(2,'0')}</span>
        </header>
        {stats.months.length === 0 ? (
          <p className="km-empty">Sem dados ainda.</p>
        ) : (
          <ul className="km-month-rows">
            {stats.months.map((m) => {
              const [mm, yy] = m.key.split('/');
              return (
                <li key={m.key} className="km-month-row">
                  <div>
                    <div className="km-month-title">{MONTH_NAMES[parseInt(mm)-1]} <span className="km-mono km-muted">·{yy}</span></div>
                    <div className="km-mono km-mono-tiny km-muted">{m.count} VIAGENS · {numFmt.format(m.km)} KM · {m.daysCount} DIAS</div>
                  </div>
                  <div className="km-month-money">{brlFmt.format((m.km * RATE) + m.toll + m.parking)}</div>
                </li>
              );
            })}
          </ul>
        )}
      </section>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SPLASH SCREEN
// ═══════════════════════════════════════════════════════════════════════════
function Splash({ exiting }) {
  return (
    <div className={`km-splash ${exiting ? 'km-splash--exit' : ''}`}>
      <div className="km-splash-grid" />
      <div className="km-splash-content">
        <div className="km-splash-mono">[ INIT · {APP_VERSION} ]</div>
        <h1 className="km-splash-title">
          <span>DRIVE</span>
          <span className="km-splash-italic">LOG</span>
        </h1>
        <div className="km-splash-bar"><div className="km-splash-bar-fill" /></div>
        <div className="km-splash-mono km-splash-mono--small">SYS · BOOT · OK</div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════════════
export default function KmTracker() {
  const [origin, setOrigin] = useState({ address: '', lat: null, lng: null });
  const [destination, setDestination] = useState({ address: '', lat: null, lng: null });
  const [distance, setDistance] = useState(null);
  const [distanceLabel, setDistanceLabel] = useState('');
  const [routeGeometry, setRouteGeometry] = useState(null);
  const [trips, setTrips] = useState([]);
  const [loading, setLoading] = useState({ origin: false, destination: false, distance: false });
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [activeTab, setActiveTab] = useState('history');
  const [time, setTime] = useState('');
  const [theme, setThemeState] = useState('light');
  const [soundOn, setSoundOn] = useState(true);
  const [hapticOn, setHapticOn] = useState(true);
  const [splashExiting, setSplashExiting] = useState(false);
  const [splashGone, setSplashGone] = useState(false);
  const [topLoading, setTopLoading] = useState(false);
  const [receipts, setReceipts] = useState([]);
  const [expenseTab, setExpenseTab] = useState('pedagio');
  const [scanning, setScanning] = useState(false);
  const [showApiKeyModal, setShowApiKeyModal] = useState(false);
  const [apiKeyInput, setApiKeyInput] = useState('');
  const [pendingReceipt, setPendingReceipt] = useState(null); // {data, imageUrl, base64}
  const [receiptPreviewId, setReceiptPreviewId] = useState(null);
  const [receiptPreviewUrl, setReceiptPreviewUrl] = useState(null);
  const fileInputRef = useRef(null);
  const initialized = useRef(false);
  const feedback = useFeedback();

  // ─── EDIT TRIP STATE ────────────────────────────────────────────────────
  const [editingTripId, setEditingTripId] = useState(null);
  const [editOrigin, setEditOrigin] = useState('');
  const [editDestination, setEditDestination] = useState('');
  const [editLoading, setEditLoading] = useState(false);

  // ─── AUTOCOMPLETE SUGGESTIONS ───────────────────────────────────────────
  const [originSuggestions, setOriginSuggestions] = useState([]);
  const [destinationSuggestions, setDestinationSuggestions] = useState([]);

  // ─── INIT: theme, sound prefs, splash ──────────────────────────────────
  useEffect(() => {
    setThemeState(getInitialTheme());
    setSoundOn(localStorage.getItem(SOUND_KEY) !== 'off');
    setHapticOn(localStorage.getItem(HAPTIC_KEY) !== 'off');

    // Splash exit timing
    const t1 = setTimeout(() => setSplashExiting(true), 1100);
    const t2 = setTimeout(() => setSplashGone(true), 1700);
    return () => { clearTimeout(t1); clearTimeout(t2); };
  }, []);

  // Live clock
  useEffect(() => {
    const update = () => {
      const d = new Date();
      setTime(`${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`);
    };
    update();
    const id = setInterval(update, 1000);
    return () => clearInterval(id);
  }, []);

  // Load trips
  useEffect(() => {
    const loaded = loadTrips();
    if (loaded && Array.isArray(loaded)) setTrips(loaded);
    else { setTrips(SEED_TRIPS); saveTrips(SEED_TRIPS); }
  }, []);

  // Load receipts
  useEffect(() => {
    setReceipts(loadReceiptsMeta());
  }, []);

  // Persist trips
  useEffect(() => { if (trips.length > 0 || loadTrips()) saveTrips(trips); }, [trips]);

  // Persist receipts
  useEffect(() => { saveReceiptsMeta(receipts); }, [receipts]);

  // Inject fonts + global CSS
  useEffect(() => {
    if (initialized.current) return;
    initialized.current = true;
    const fonts = document.createElement('link');
    fonts.href = 'https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600;700;800;900&family=Geist+Mono:wght@400;500;700&family=Instrument+Serif:ital@0;1&display=swap';
    fonts.rel = 'stylesheet';
    document.head.appendChild(fonts);

    const style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);
  }, []);

  // Distance calculation
  useEffect(() => {
    if (origin.lat == null || destination.lat == null) {
      setDistance(null); setDistanceLabel(''); setRouteGeometry(null); return;
    }
    let cancelled = false;
    (async () => {
      setLoading(s => ({ ...s, distance: true }));
      try {
        const r = await fetch(`https://router.project-osrm.org/route/v1/driving/${origin.lng},${origin.lat};${destination.lng},${destination.lat}?overview=simplified&geometries=geojson`);
        const data = await r.json();
        if (cancelled) return;
        if (data.routes?.length > 0) {
          const route = data.routes[0];
          setDistance(route.distance / 1000);
          setDistanceLabel('rota de carro');
          setRouteGeometry(route.geometry?.coordinates || null);
        } else throw new Error();
      } catch {
        if (cancelled) return;
        setDistance(haversineKm(origin.lat, origin.lng, destination.lat, destination.lng));
        setDistanceLabel('linha reta · rota indisponível');
        setRouteGeometry([[origin.lng, origin.lat], [destination.lng, destination.lat]]);
      } finally {
        if (!cancelled) setLoading(s => ({ ...s, distance: false }));
      }
    })();
    return () => { cancelled = true; };
  }, [origin.lat, origin.lng, destination.lat, destination.lng]);

  // ─── AUTOCOMPLETE (INTELIGÊNCIA 2.0) ────────────────────────────────────
  useEffect(() => {
    // Se já tem lat/lng, não precisa de sugestões
    if (!origin.address || origin.lat != null || origin.address.length < 5) {
      setOriginSuggestions([]); return;
    }
    const apiKey = localStorage.getItem(GEMINI_KEY);
    if (!apiKey) return;

    const timer = setTimeout(async () => {
      setLoading(prev => ({ ...prev, origin: true }));
      const list = await resolvePlaceWithGemini(origin.address, apiKey);
      setOriginSuggestions(list || []);
      setLoading(prev => ({ ...prev, origin: false }));
    }, 1000);

    return () => clearTimeout(timer);
  }, [origin.address, origin.lat]);

  useEffect(() => {
    // Se já tem lat/lng, não precisa de sugestões
    if (!destination.address || destination.lat != null || destination.address.length < 5) {
      setDestinationSuggestions([]); return;
    }
    const apiKey = localStorage.getItem(GEMINI_KEY);
    if (!apiKey) return;

    const timer = setTimeout(async () => {
      setLoading(prev => ({ ...prev, destination: true }));
      const list = await resolvePlaceWithGemini(destination.address, apiKey);
      setDestinationSuggestions(list || []);
      setLoading(prev => ({ ...prev, destination: false }));
    }, 1000);

    return () => clearTimeout(timer);
  }, [destination.address, destination.lat]);

  // ─── HANDLERS ───────────────────────────────────────────────────────────
  async function capture(which) {
    setError(null); setLoading(s => ({ ...s, [which]: true }));
    setTopLoading(true);
    feedback('tap');
    try {
      if (!navigator.geolocation) throw new Error('Geolocalização não disponível neste navegador');
      const pos = await new Promise((res, rej) => navigator.geolocation.getCurrentPosition(
        p => res({ lat: p.coords.latitude, lng: p.coords.longitude }),
        e => rej(new Error(
          e.code === 1 ? 'Permissão de localização negada. Habilite o GPS para este site.' :
          e.code === 2 ? 'Localização indisponível. Verifique se o GPS está ligado.' :
          e.code === 3 ? 'Tempo esgotado. Tente novamente.' : 'Erro ao obter localização'
        )),
        { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
      ));
      const result = await getAddressWithNumber(pos.lat, pos.lng);
      const address = formatBrazilianAddress(result.data, result.number, result.isApprox);
      (which === 'origin' ? setOrigin : setDestination)({ address, lat: pos.lat, lng: pos.lng });
      feedback('success');
      if (result.number && result.isApprox) {
        setSuccess('Número aproximado · pode ajustar');
        setTimeout(() => setSuccess(null), 2600);
      } else if (!result.number) {
        setSuccess('Sem número · adicione manualmente');
        setTimeout(() => setSuccess(null), 2600);
      }
    } catch (e) {
      setError(e.message); feedback('error');
    } finally {
      setLoading(s => ({ ...s, [which]: false }));
      setTopLoading(false);
    }
  }

  function clearLocation(which) {
    feedback('tap');
    (which === 'origin' ? setOrigin : setDestination)({ address: '', lat: null, lng: null });
  }

  function selectSuggestion(which, s) {
    if (which === 'origin') {
      setOrigin({ address: s.address, lat: s.lat, lng: s.lng });
      setOriginSuggestions([]);
    } else {
      setDestination({ address: s.address, lat: s.lat, lng: s.lng });
      setDestinationSuggestions([]);
    }
    feedback('success');
  }

  function handleSave() {
    if (!origin.address.trim() || !destination.address.trim()) {
      setError('Preencha origem e destino'); feedback('error'); return;
    }
    const d = new Date();
    const newTrip = {
      id: 't' + Date.now(),
      date: `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`,
      origin: origin.address.trim(),
      destination: destination.address.trim(),
      km: distance != null ? Number(distance.toFixed(2)) : null,
      kmLabel: distanceLabel,
      geometry: routeGeometry,
    };
    setTrips(prev => [newTrip, ...prev]);
    setOrigin({ address:'', lat:null, lng:null });
    setDestination({ address:'', lat:null, lng:null });
    setDistance(null); setDistanceLabel(''); setRouteGeometry(null); setError(null);
    setSuccess('Viagem registrada'); setTimeout(() => setSuccess(null), 2200);
    feedback('success');
  }

  function deleteTrip(id) {
    if (window.confirm('Remover esta viagem?')) {
      setTrips(prev => prev.filter(t => t.id !== id));
      if (editingTripId === id) setEditingTripId(null);
      feedback('swoosh');
    }
  }

  // ─── EDIT TRIP HANDLERS ─────────────────────────────────────────────────
  function startEditTrip(trip) {
    setEditingTripId(trip.id);
    setEditOrigin(trip.origin);
    setEditDestination(trip.destination);
    feedback('tap');
  }

  function cancelEditTrip() {
    setEditingTripId(null);
    setEditOrigin('');
    setEditDestination('');
    feedback('tap');
  }

  async function saveEditTrip() {
    if (!editOrigin.trim() || !editDestination.trim()) {
      setError('Preencha origem e destino'); feedback('error'); return;
    }
    setEditLoading(true);
    setTopLoading(true);
    feedback('tap');
    try {
      // Geocode origin
      const originGeo = await geocodeAddress(editOrigin.trim());
      // Geocode destination
      const destGeo = await geocodeAddress(editDestination.trim());

      let km = null;
      let kmLabel = '';
      let geometry = null;

      if (originGeo && destGeo) {
        // Try OSRM route
        try {
          const r = await fetch(`https://router.project-osrm.org/route/v1/driving/${originGeo.lng},${originGeo.lat};${destGeo.lng},${destGeo.lat}?overview=simplified&geometries=geojson`);
          const data = await r.json();
          if (data.routes?.length > 0) {
            km = Number((data.routes[0].distance / 1000).toFixed(2));
            kmLabel = 'rota de carro';
            geometry = data.routes[0].geometry?.coordinates || null;
          }
        } catch {}

        if (km == null) {
          km = Number(haversineKm(originGeo.lat, originGeo.lng, destGeo.lat, destGeo.lng).toFixed(2));
          kmLabel = 'linha reta · rota indisponível';
          geometry = [[originGeo.lng, originGeo.lat], [destGeo.lng, destGeo.lat]];
        }
      }

      setTrips(prev => prev.map(t => {
        if (t.id !== editingTripId) return t;
        return {
          ...t,
          origin: editOrigin.trim(),
          destination: editDestination.trim(),
          km,
          kmLabel,
          geometry,
        };
      }));

      setEditingTripId(null);
      setEditOrigin('');
      setEditDestination('');
      setSuccess('Viagem atualizada!'); setTimeout(() => setSuccess(null), 2200);
      feedback('success');
    } catch (e) {
      setError(e.message || 'Erro ao recalcular rota');
      feedback('error');
    } finally {
      setEditLoading(false);
      setTopLoading(false);
    }
  }

  // ─── RECEIPT HANDLERS ──────────────────────────────────────────────────
  async function handleReceiptPhoto(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const apiKey = localStorage.getItem(GEMINI_KEY);
    if (!apiKey) {
      setShowApiKeyModal(true);
      return;
    }
    setScanning(true);
    setError(null);
    setTopLoading(true);
    feedback('tap');
    try {
      // 1. Carrega imagem original na tela
      const originalImageUrl = URL.createObjectURL(file);
      const img = new Image();
      img.src = originalImageUrl;
      await new Promise(r => img.onload = r);
      
      // 2. OpenCV Canny Edge Detection & Warp Perspective
      const cropped = await autoCropDocument(img);
      const finalBlob = cropped.blob;
      const finalBase64 = cropped.dataUrl.split(',')[1];
      const finalImageUrl = URL.createObjectURL(finalBlob);
      URL.revokeObjectURL(originalImageUrl);

      // 3. Envia a imagem perfeita (crop) pro Gemini extrair os dados
      const data = await analyzeReceiptWithGemini(finalBase64, apiKey);
      
      setPendingReceipt({ data, imageUrl: finalImageUrl, base64: finalBase64, blob: finalBlob });
      feedback('success');
    } catch (err) {
      setError(err.message || 'Erro ao analisar comprovante');
      feedback('error');
    } finally {
      setScanning(false);
      setTopLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }

  async function confirmReceipt() {
    if (!pendingReceipt) return;
    const id = 'r' + Date.now();
    const meta = {
      id,
      ...pendingReceipt.data,
      createdAt: new Date().toISOString(),
    };
    try {
      await saveReceiptImage(id, pendingReceipt.blob);
    } catch { /* IndexedDB fail — continue without image */ }
    setReceipts(prev => [meta, ...prev]);
    setPendingReceipt(null);
    setSuccess('Comprovante salvo!');
    setTimeout(() => setSuccess(null), 2200);
    feedback('success');
  }

  function cancelReceipt() {
    if (pendingReceipt?.imageUrl) URL.revokeObjectURL(pendingReceipt.imageUrl);
    setPendingReceipt(null);
  }

  async function deleteReceipt(id) {
    if (!window.confirm('Remover este comprovante?')) return;
    try { await deleteReceiptImage(id); } catch {}
    setReceipts(prev => prev.filter(r => r.id !== id));
    feedback('swoosh');
  }

  async function downloadReceipt(id) {
    try {
      const blob = await getReceiptImage(id);
      if (!blob) { setError('Imagem não encontrada'); return; }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `comprovante_${id}.jpg`;
      a.click();
      URL.revokeObjectURL(url);
      feedback('success');
    } catch {
      setError('Erro ao baixar comprovante');
    }
  }

  async function downloadReceiptPdf(id) {
    try {
      const blob = await getReceiptImage(id);
      if (!blob) { setError('Imagem não encontrada'); return; }
      const meta = receipts.find(r => r.id === id) || {};
      
      const dataUrl = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      });
      
      const pdf = new jsPDF();
      const img = new Image();
      img.src = dataUrl;
      await new Promise(res => img.onload = res);
      
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const imgRatio = img.width / img.height;
      
      let renderWidth = pageWidth - 20;
      let renderHeight = renderWidth / imgRatio;
      
      if (renderHeight > pageHeight - 60) {
         renderHeight = pageHeight - 60;
         renderWidth = renderHeight * imgRatio;
      }
      
      pdf.setFontSize(16);
      pdf.text(`Comprovante: ${meta.concessionaria || meta.tipo || 'Recibo'}`, 10, 20);
      pdf.setFontSize(11);
      pdf.text(`Data: ${meta.dataHora || 'N/A'} - Valor: R$ ${Number(meta.valor||0).toFixed(2)}`, 10, 28);
      
      pdf.addImage(dataUrl, 'JPEG', 10, 40, renderWidth, renderHeight);
      pdf.save(`comprovante_${id}.pdf`);
      feedback('success');
    } catch (e) {
      setError('Erro ao gerar PDF');
      console.error(e);
    }
  }

  async function viewReceipt(id) {
    try {
      const blob = await getReceiptImage(id);
      if (!blob) { setError('Imagem não encontrada'); return; }
      const url = URL.createObjectURL(blob);
      setReceiptPreviewId(id);
      setReceiptPreviewUrl(url);
    } catch {
      setError('Erro ao carregar imagem');
    }
  }

  function closeReceiptPreview() {
    if (receiptPreviewUrl) URL.revokeObjectURL(receiptPreviewUrl);
    setReceiptPreviewId(null);
    setReceiptPreviewUrl(null);
  }

  function handleSaveApiKey() {
    if (!apiKeyInput.trim()) return;
    localStorage.setItem(GEMINI_KEY, apiKeyInput.trim());
    setShowApiKeyModal(false);
    setApiKeyInput('');
    setSuccess('API Key salva!');
    setTimeout(() => setSuccess(null), 2200);
    feedback('success');
  }

  function openScanWithApiCheck() {
    const apiKey = localStorage.getItem(GEMINI_KEY);
    if (!apiKey) {
      setShowApiKeyModal(true);
      return;
    }
    fileInputRef.current?.click();
  }

  function exportToExcel() {
    if (!trips.length && !receipts.length) { setError('Nenhum dado para exportar'); feedback('error'); return; }
    const wb = XLSX.utils.book_new();

    // ABA 1: Viagens
    const sortedTrips = [...trips].sort((a, b) => {
      const parse = s => { const [d,m,y] = s.split('/'); return new Date(y,m-1,d); };
      return parse(a.date) - parse(b.date);
    });
    const tripsData = [['Data','Origem','Destino','KM','Valor (R$)'], ...sortedTrips.map(t => [t.date, t.origin, t.destination, t.km || '', t.km ? (t.km * RATE).toFixed(2) : ''])];
    const wsTrips = XLSX.utils.aoa_to_sheet(tripsData);
    wsTrips['!cols'] = [{ wch:12 },{ wch:55 },{ wch:55 },{ wch:10 },{ wch:12 }];
    XLSX.utils.book_append_sheet(wb, wsTrips, 'kmadicional');

    // ABA 2: Pedágios
    const pedData = [['Data/Hora','Concessionária','Praça','Via','Placa','Valor (R$)','Recibo'], ...pedagios.map(r => [r.dataHora||'', r.concessionaria||'', r.praca||'', r.via||'', r.placa||'', r.valor||0, r.recibo||''])];
    const wsPed = XLSX.utils.aoa_to_sheet(pedData);
    wsPed['!cols'] = [{ wch:20 },{ wch:30 },{ wch:30 },{ wch:25 },{ wch:10 },{ wch:12 },{ wch:20 }];
    XLSX.utils.book_append_sheet(wb, wsPed, 'Pedágios');

    // ABA 3: Estacionamentos
    const estData = [['Data/Hora','Estabelecimento','Local','Endereço','Placa','Valor (R$)','Recibo'], ...estacionamentos.map(r => [r.dataHora||'', r.concessionaria||'', r.praca||'', r.via||'', r.placa||'', r.valor||0, r.recibo||''])];
    const wsEst = XLSX.utils.aoa_to_sheet(estData);
    wsEst['!cols'] = [{ wch:20 },{ wch:30 },{ wch:30 },{ wch:25 },{ wch:10 },{ wch:12 },{ wch:20 }];
    XLSX.utils.book_append_sheet(wb, wsEst, 'Estacionamentos');

    XLSX.writeFile(wb, 'relatorio_completo_km.xlsx');
    setSuccess('Planilha exportada com 3 abas!'); setTimeout(() => setSuccess(null), 2200);
    feedback('success');
  }

  async function generateFullReport() {
    setTopLoading(true);
    feedback('tap');
    try {
      const pdf = new jsPDF();
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const now = new Date();
      const periodo = `${MONTH_FULL[now.getMonth()]} ${now.getFullYear()}`;

      // ── CAPA ──
      pdf.setFillColor(10, 10, 11);
      pdf.rect(0, 0, pageW, pageH, 'F');
      pdf.setTextColor(237, 237, 237);
      pdf.setFontSize(36);
      pdf.text('DRIVE LOG', pageW / 2, 80, { align: 'center' });
      pdf.setFontSize(14);
      pdf.text('RELATÓRIO DE REEMBOLSO', pageW / 2, 95, { align: 'center' });
      pdf.setFontSize(11);
      pdf.setTextColor(161, 161, 170);
      pdf.text(`Período: ${periodo}`, pageW / 2, 115, { align: 'center' });
      pdf.text(`Gerado em: ${now.toLocaleDateString('pt-BR')} às ${now.toLocaleTimeString('pt-BR')}`, pageW / 2, 125, { align: 'center' });
      pdf.setTextColor(230, 26, 39);
      pdf.setFontSize(28);
      pdf.text(brlFmt.format(totalEarning + totalDespesas), pageW / 2, 160, { align: 'center' });
      pdf.setFontSize(10);
      pdf.setTextColor(161, 161, 170);
      pdf.text('VALOR TOTAL A REEMBOLSAR (KM + DESPESAS)', pageW / 2, 172, { align: 'center' });

      // ── RESUMO KM ──
      pdf.addPage();
      pdf.setTextColor(10, 10, 11);
      pdf.setFontSize(18);
      pdf.text('Resumo de Quilometragem', 10, 20);
      pdf.setFontSize(10);
      pdf.setTextColor(82, 82, 91);
      pdf.text(`Total: ${numFmt.format(totalKm)} KM × R$ 1,14 = ${brlFmt.format(totalEarning)}`, 10, 30);

      let y = 42;
      pdf.setFontSize(8);
      pdf.setTextColor(120, 120, 120);
      pdf.text('DATA', 10, y);
      pdf.text('ORIGEM', 30, y);
      pdf.text('DESTINO', 110, y);
      pdf.text('KM', 185, y, { align: 'right' });
      pdf.text('R$', 200, y, { align: 'right' });
      y += 6;
      pdf.setDrawColor(200, 200, 200);
      pdf.line(10, y - 2, 200, y - 2);

      pdf.setTextColor(10, 10, 11);
      for (const t of trips) {
        if (y > pageH - 20) { pdf.addPage(); y = 20; }
        pdf.setFontSize(7.5);
        pdf.text(t.date, 10, y);
        pdf.text((t.origin || '').substring(0, 45), 30, y);
        pdf.text((t.destination || '').substring(0, 45), 110, y);
        pdf.text(t.km != null ? numFmt.format(t.km) : '—', 185, y, { align: 'right' });
        pdf.text(t.km != null ? brlFmt.format(t.km * RATE) : '—', 200, y, { align: 'right' });
        y += 5;
      }

      // ── RESUMO DESPESAS ──
      if (receipts.length > 0) {
        pdf.addPage();
        pdf.setFontSize(18);
        pdf.setTextColor(10, 10, 11);
        pdf.text('Resumo de Despesas', 10, 20);
        pdf.setFontSize(10);
        pdf.setTextColor(82, 82, 91);
        pdf.text(`Pedágios: ${brlFmt.format(totalPedagios)} | Estacionamentos: ${brlFmt.format(totalEstacionamentos)} | Total: ${brlFmt.format(totalDespesas)}`, 10, 30);

        y = 42;
        pdf.setFontSize(8);
        pdf.setTextColor(120, 120, 120);
        pdf.text('DATA', 10, y);
        pdf.text('TIPO', 40, y);
        pdf.text('CONCESSÃO/LOCAL', 65, y);
        pdf.text('PRAÇA', 130, y);
        pdf.text('VALOR', 200, y, { align: 'right' });
        y += 6;
        pdf.line(10, y - 2, 200, y - 2);

        pdf.setTextColor(10, 10, 11);
        for (const r of receipts) {
          if (y > pageH - 20) { pdf.addPage(); y = 20; }
          pdf.setFontSize(7.5);
          pdf.text(r.dataHora || '—', 10, y);
          pdf.text(r.tipo === 'pedagio' ? 'Pedágio' : 'Estacion.', 40, y);
          pdf.text((r.concessionaria || '—').substring(0, 35), 65, y);
          pdf.text((r.praca || '—').substring(0, 35), 130, y);
          pdf.text(brlFmt.format(r.valor || 0), 200, y, { align: 'right' });
          y += 5;
        }

        // ── COMPROVANTES (1 por página) ──
        for (const r of receipts) {
          try {
            const blob = await getReceiptImage(r.id);
            if (!blob) continue;
            const dataUrl = await new Promise(res => {
              const reader = new FileReader();
              reader.onload = () => res(reader.result);
              reader.readAsDataURL(blob);
            });
            pdf.addPage();
            pdf.setFontSize(14);
            pdf.setTextColor(10, 10, 11);
            pdf.text(`Comprovante: ${r.concessionaria || r.tipo || 'Recibo'}`, 10, 18);
            pdf.setFontSize(9);
            pdf.setTextColor(82, 82, 91);
            pdf.text(`Data: ${r.dataHora || 'N/A'} | Valor: ${brlFmt.format(r.valor || 0)} | Via: ${r.via || 'N/A'}`, 10, 26);
            if (r.recibo) pdf.text(`Recibo: ${r.recibo}`, 10, 32);

            const img = new Image();
            img.src = dataUrl;
            await new Promise(res => { img.onload = res; img.onerror = res; });
            const imgRatio = img.width / img.height;
            let rw = pageW - 20, rh = rw / imgRatio;
            if (rh > pageH - 50) { rh = pageH - 50; rw = rh * imgRatio; }
            pdf.addImage(dataUrl, 'JPEG', 10, 38, rw, rh);
          } catch { /* skip */ }
        }
      }

      // ── TOTALIZAÇÃO FINAL ──
      pdf.addPage();
      pdf.setFillColor(10, 10, 11);
      pdf.rect(0, 0, pageW, pageH, 'F');
      pdf.setTextColor(237, 237, 237);
      pdf.setFontSize(14);
      pdf.text('TOTALIZAÇÃO', pageW / 2, 60, { align: 'center' });

      pdf.setFontSize(11);
      pdf.setTextColor(161, 161, 170);
      pdf.text(`Quilometragem: ${numFmt.format(totalKm)} KM`, pageW / 2, 85, { align: 'center' });
      pdf.text(`Valor KM: ${brlFmt.format(totalEarning)}`, pageW / 2, 95, { align: 'center' });
      pdf.text(`Pedágios: ${brlFmt.format(totalPedagios)}`, pageW / 2, 105, { align: 'center' });
      pdf.text(`Estacionamentos: ${brlFmt.format(totalEstacionamentos)}`, pageW / 2, 115, { align: 'center' });

      pdf.setTextColor(230, 26, 39);
      pdf.setFontSize(32);
      pdf.text(brlFmt.format(totalEarning + totalDespesas), pageW / 2, 150, { align: 'center' });
      pdf.setFontSize(10);
      pdf.setTextColor(161, 161, 170);
      pdf.text('TOTAL A REEMBOLSAR', pageW / 2, 162, { align: 'center' });

      pdf.save(`relatorio_reembolso_${now.getMonth()+1}_${now.getFullYear()}.pdf`);
      setSuccess('Relatório PDF gerado!'); setTimeout(() => setSuccess(null), 2200);
      feedback('success');
    } catch (e) {
      console.error(e);
      setError('Erro ao gerar relatório');
      feedback('error');
    } finally {
      setTopLoading(false);
    }
  }

  function handleToggleTheme() {
    feedback('tap');
    const next = toggleTheme(theme);
    setThemeState(next);
  }
  function handleToggleSound() {
    const next = !soundOn;
    setSoundOn(next);
    localStorage.setItem(SOUND_KEY, next ? 'on' : 'off');
    if (next) feedback('tap');
  }
  function handleToggleHaptic() {
    const next = !hapticOn;
    setHapticOn(next);
    localStorage.setItem(HAPTIC_KEY, next ? 'on' : 'off');
    if (next && 'vibrate' in navigator) navigator.vibrate(15);
  }
  function handleTabSwitch(t) {
    if (t === activeTab) return;
    feedback('swoosh');
    if (document.startViewTransition) {
      document.startViewTransition(() => setActiveTab(t));
    } else {
      setActiveTab(t);
    }
  }

  // ─── COMPUTED ───────────────────────────────────────────────────────────
  const totalKm = trips.reduce((s, t) => s + (t.km || 0), 0);
  const kmEarning = totalKm * RATE;
  const tripsWithKm = trips.filter(t => t.km != null).length;
  const canSave = origin.address.trim() && destination.address.trim();

  const pedagios = receipts.filter(r => r.tipo === 'pedagio');
  const estacionamentos = receipts.filter(r => r.tipo === 'estacionamento');
  const totalPedagios = pedagios.reduce((s, r) => s + (r.valor || 0), 0);
  const totalEstacionamentos = estacionamentos.reduce((s, r) => s + (r.valor || 0), 0);
  const totalDespesas = totalPedagios + totalEstacionamentos;
  const totalEarning = kmEarning + totalDespesas;

  async function backfillMissingDfes() {
    const apiKey = localStorage.getItem(GEMINI_KEY);
    if (!apiKey) { setShowApiKeyModal(true); return; }

    const pending = receipts.filter(r => r.tipo === 'pedagio' && !r.dfe);
    if (pending.length === 0) {
      setSuccess('Todos os recibos já possuem DFE');
      setTimeout(() => setSuccess(null), 2000);
      return;
    }

    setTopLoading(true);
    let count = 0;
    for (const r of pending) {
      try {
        const blob = await getReceiptImage(r.id);
        if (!blob) continue;
        const base64 = await new Promise((res) => {
          const reader = new FileReader();
          reader.onload = () => res(reader.result.split(',')[1]);
          reader.readAsDataURL(blob);
        });
        const data = await analyzeReceiptWithGemini(base64, apiKey);
        if (data && data.dfe) {
          setReceipts(prev => prev.map(item => item.id === r.id ? { ...item, dfe: data.dfe } : item));
          count++;
        }
      } catch (e) {
        console.error('Erro no backfill:', e);
      }
    }
    setTopLoading(false);
    setSuccess(`${count} recibos atualizados com DFE`);
    setTimeout(() => setSuccess(null), 3000);
    feedback('success');
  }

  const marqueeText = ' · DRIVE LOG · COCA-COLA BR · FROTA · ' + numFmt.format(totalKm) + ' KM · ' + brlFmt.format(totalEarning) + ' · ' + String(trips.length).padStart(3,'0') + ' VIAGENS' + (totalDespesas > 0 ? ' · ' + brlFmt.format(totalDespesas) + ' DESPESAS' : '');

  return (
    <div className="km-app" data-theme={theme}>
      {!splashGone && <Splash exiting={splashExiting} />}

      {topLoading && <div className="km-toploader" />}

      {/* STATUS BAR */}
      <div className="km-statusbar">
        <div className="km-statusbar-inner">
          <span className="km-mono km-mono-tiny">
            <span className="km-status-dot" /> SYS · ONLINE
          </span>
          <span className="km-mono km-mono-tiny km-muted">{APP_VERSION} · DRIVE LOG</span>
          <span className="km-mono km-mono-tiny">{time}</span>
        </div>
      </div>

      {/* HEADER */}
      <header className="km-header">
        <div className="km-aurora-bg" />
        <div className="km-grid-overlay" />

        <div className="km-header-top">
          <div>
            <div className="km-mono km-mono-tiny km-muted">[ {APP_VERSION} ] · MOTORISTA</div>
            <h1 className="km-h1">
              <span>CONTROLE</span>
              <span className="km-h1-italic">de KM</span>
            </h1>
          </div>
          <div className="km-header-actions">
            <button onClick={handleToggleTheme} className="km-icon-btn" aria-label="Trocar tema">
              {theme === 'dark' ? <Sun size={16} strokeWidth={2.4} /> : <Moon size={16} strokeWidth={2.4} />}
            </button>
            <button onClick={handleToggleSound} className="km-icon-btn" aria-label="Som">
              {soundOn ? <Volume2 size={16} strokeWidth={2.4} /> : <VolumeX size={16} strokeWidth={2.4} />}
            </button>
            <button onClick={handleToggleHaptic} className="km-icon-btn" aria-label="Vibração">
              <Vibrate size={16} strokeWidth={2.4} style={{ opacity: hapticOn ? 1 : 0.35 }} />
            </button>
          </div>
        </div>

        {/* HERO METRIC EDITORIAL */}
        <div className="km-hero">
          <div className="km-mono km-mono-label">
            <span className="km-line" />
            VALOR A RECEBER
          </div>
          <div className="km-hero-value">
            <BigCurrency value={totalEarning} size="hero" />
          </div>
          <div className="km-hero-meta-row">
            <span className="km-pill">
              <span className="km-status-dot km-status-dot--green" />
              <span className="km-mono km-mono-tiny">{numFmt.format(totalKm)} KM TOTAIS</span>
            </span>
            <span className="km-mono km-mono-tiny km-muted">
              {brlFmt.format(kmEarning)} KM + {brlFmt.format(totalDespesas)} DESP
            </span>
          </div>
        </div>

        {/* WAVE */}
        <svg viewBox="0 0 400 60" preserveAspectRatio="none" className="km-wave">
          <path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25 L400,60 L0,60 Z" fill="var(--bg-base)" />
          <path d="M0,40 C60,10 120,55 200,30 C280,5 340,50 400,25" fill="none" stroke="rgba(255,255,255,0.4)" strokeWidth="1" />
        </svg>
      </header>

      {/* CONTENT */}
      <main className="km-main">
        {/* ALERTS */}
        {error && (
          <div className="km-alert km-alert--err km-fade-in">
            <AlertCircle size={18} strokeWidth={2.4} />
            <p>{error}</p>
            <button onClick={() => setError(null)}><X size={14} /></button>
          </div>
        )}
        {success && (
          <div className="km-alert km-alert--ok km-fade-in">
            <span className="km-status-dot km-status-dot--green" />
            <p>{success}</p>
          </div>
        )}

        {/* NEW TRIP */}
        <section className="km-card">
          <header className="km-card-head km-card-head--bordered">
            <div className="km-row-tight">
              <span className="km-redbar" />
              <h2 className="km-h2">Nova viagem</h2>
            </div>
            <span className="km-mono km-mono-tiny km-muted">001/REC</span>
          </header>

          {/* Origin */}
          <div className="km-field">
            <div className="km-field-head">
              <label className="km-mono km-mono-label">
                <span className="km-dot km-dot--red" />· ORIGEM
              </label>
              {origin.address && (
                <button onClick={() => clearLocation('origin')} className="km-textbtn">CLEAR</button>
              )}
            </div>
            <div style={{ position: 'relative' }}>
              <textarea value={origin.address} onChange={e => setOrigin({ address: e.target.value, lat: null, lng: null })} placeholder="Endereço de partida" rows={2} className="km-textarea" />
              <SuggestionsList suggestions={originSuggestions} onSelect={(s) => selectSuggestion('origin', s)} loading={loading.origin} />
            </div>
            <button onClick={() => capture('origin')} disabled={loading.origin} className="km-btn km-btn--red km-btn--block km-press">
              {loading.origin ? <><Loader2 size={15} className="km-spin" /> BUSCANDO…</> : <><MapPin size={15} strokeWidth={2.4} /> CAPTURAR LOCALIZAÇÃO</>}
            </button>
          </div>

          {/* Connector */}
          <div className="km-connector">
            <ArrowDown size={11} strokeWidth={2.4} />
            <div className="km-dashed" />
            <span className="km-mono km-mono-tiny">DESTINO</span>
            <div className="km-dashed" />
          </div>

          {/* Destination */}
          <div className="km-field">
            <div className="km-field-head">
              <label className="km-mono km-mono-label">
                <Flag size={11} strokeWidth={2.4} fill="currentColor" />· DESTINO
              </label>
              {destination.address && (
                <button onClick={() => clearLocation('destination')} className="km-textbtn">CLEAR</button>
              )}
            </div>
            <div style={{ position: 'relative' }}>
              <textarea value={destination.address} onChange={e => setDestination({ address: e.target.value, lat: null, lng: null })} placeholder="Endereço de chegada" rows={2} className="km-textarea" />
              <SuggestionsList suggestions={destinationSuggestions} onSelect={(s) => selectSuggestion('destination', s)} loading={loading.destination} />
            </div>
            <button onClick={() => capture('destination')} disabled={loading.destination} className="km-btn km-btn--ink km-btn--block km-press">
              {loading.destination ? <><Loader2 size={15} className="km-spin" /> BUSCANDO…</> : <><Flag size={15} strokeWidth={2.4} /> CAPTURAR LOCALIZAÇÃO</>}
            </button>
          </div>

          {/* Odometer (BRUTALIST DARK) */}
          <div className="km-odometer">
            <div className="km-odometer-glow" />
            <div className="km-odometer-grid" />
            <div className="km-odometer-info">
              <div className="km-mono km-mono-label km-odometer-label">◊ DISTÂNCIA</div>
              {distance != null && (
                <div className="km-odometer-money">{brlFmt.format(distance * RATE)}</div>
              )}
              {distanceLabel && (
                <div className="km-mono km-mono-tiny km-odometer-sublabel">{distanceLabel}</div>
              )}
            </div>
            <div className="km-odometer-display">
              {loading.distance
                ? <Loader2 size={28} className="km-spin" style={{ color: 'var(--coca-red)' }} />
                : distance != null
                  ? <div className="km-odometer-num">
                      <span>{numFmt.format(distance)}</span>
                      <span className="km-odometer-unit">KM</span>
                    </div>
                  : <div className="km-odometer-empty">--.--</div>
              }
            </div>
          </div>

          <button onClick={handleSave} disabled={!canSave} className="km-btn km-btn--save km-btn--block km-press">
            <Save size={16} strokeWidth={2.4} /> REGISTRAR VIAGEM
          </button>
        </section>

        {/* TABS */}
        <div className="km-tabs">
          <button onClick={() => handleTabSwitch('history')} className={`km-tab ${activeTab === 'history' ? 'km-tab--active' : ''}`}>
            <History size={14} strokeWidth={2.4} /> HISTÓRICO
          </button>
          <button onClick={() => handleTabSwitch('expenses')} className={`km-tab ${activeTab === 'expenses' ? 'km-tab--active' : ''}`}>
            <Receipt size={14} strokeWidth={2.4} /> DESPESAS
          </button>
          <button onClick={() => handleTabSwitch('dashboard')} className={`km-tab ${activeTab === 'dashboard' ? 'km-tab--active' : ''}`}>
            <TrendingUp size={14} strokeWidth={2.4} /> DASHBOARD
          </button>
        </div>

        {/* HIDDEN FILE INPUT FOR CAMERA */}
        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          capture="environment"
          onChange={handleReceiptPhoto}
          style={{ display: 'none' }}
        />

        {/* TAB CONTENT */}
        <div className="km-tabcontent" key={activeTab}>
          {activeTab === 'dashboard' ? (
            <Dashboard trips={trips} receipts={receipts} />
          ) : activeTab === 'expenses' ? (
            <div className="km-fade-up">
              {/* EXPENSE SUB-TABS */}
              <div className="km-expense-subtabs">
                <button onClick={() => { setExpenseTab('pedagio'); feedback('tap'); }} className={`km-expense-subtab ${expenseTab === 'pedagio' ? 'km-expense-subtab--active' : ''}`}>
                  <Receipt size={13} strokeWidth={2.4} /> PEDÁGIOS
                  {pedagios.length > 0 && <span className="km-expense-badge">{pedagios.length}</span>}
                </button>
                <button onClick={() => { setExpenseTab('estacionamento'); feedback('tap'); }} className={`km-expense-subtab ${expenseTab === 'estacionamento' ? 'km-expense-subtab--active' : ''}`}>
                  <ParkingCircle size={13} strokeWidth={2.4} /> ESTACIONAMENTOS
                  {estacionamentos.length > 0 && <span className="km-expense-badge">{estacionamentos.length}</span>}
                </button>
              </div>

              {/* SCAN BUTTON */}
              <button onClick={openScanWithApiCheck} disabled={scanning} className="km-btn km-btn--red km-btn--block km-press" style={{ marginBottom: '14px' }}>
                {scanning ? <><Loader2 size={16} className="km-spin" /> ANALISANDO COM IA…</> : <><Camera size={16} strokeWidth={2.4} /> ESCANEAR COMPROVANTE</>}
              </button>

              {/* API KEY CONFIG */}
              <div style={{ display: 'flex', gap: '8px', marginBottom: '14px' }}>
                <button onClick={() => { setApiKeyInput(localStorage.getItem(GEMINI_KEY) || ''); setShowApiKeyModal(true); }} className="km-textbtn" style={{ fontSize: '10px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <Settings size={11} /> CONFIGURAR API KEY
                </button>
                <button onClick={backfillMissingDfes} className="km-textbtn" style={{ fontSize: '10px', display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--coca-red)' }}>
                  <Check size={11} /> ATUALIZAR DFES ANTIGOS
                </button>
              </div>

              {/* TOTAL CARD */}
              <section className="km-card" style={{ marginBottom: '14px' }}>
                <header className="km-card-head">
                  <span className="km-mono km-mono-label">◊ TOTAL A REEMBOLSAR · {expenseTab === 'pedagio' ? 'PEDÁGIOS' : 'ESTACIONAMENTOS'}</span>
                </header>
                <div style={{ marginTop: '8px' }}>
                  <BigCurrency value={expenseTab === 'pedagio' ? totalPedagios : totalEstacionamentos} size="md" />
                </div>
                {totalDespesas > 0 && (
                  <div className="km-mono km-mono-tiny km-muted" style={{ marginTop: '6px' }}>
                    TOTAL GERAL (PED + EST): {brlFmt.format(totalDespesas)}
                  </div>
                )}
              </section>

              {/* RECEIPT LIST */}
              <section className="km-card">
                <header className="km-card-head km-card-head--bordered">
                  <div className="km-row-tight">
                    <span className="km-redbar" />
                    <h2 className="km-h2">{expenseTab === 'pedagio' ? 'Pedágios' : 'Estacionamentos'}</h2>
                    <span className="km-mono km-mono-tiny km-pill-dark">{String(expenseTab === 'pedagio' ? pedagios.length : estacionamentos.length).padStart(2,'0')}</span>
                  </div>
                </header>

                {(expenseTab === 'pedagio' ? pedagios : estacionamentos).length === 0 ? (
                  <p className="km-empty">Nenhum comprovante registrado. Escaneie um {expenseTab === 'pedagio' ? 'pedágio' : 'estacionamento'}.</p>
                ) : (
                  <ul className="km-trip-list">
                    {(expenseTab === 'pedagio' ? pedagios : estacionamentos).map((r, idx) => (
                      <li key={r.id} className="km-trip">
                        <div className="km-trip-head">
                          <span className="km-mono km-mono-tiny">
                            #{String((expenseTab === 'pedagio' ? pedagios : estacionamentos).length - idx).padStart(3,'0')} <span className="km-muted">·</span> {r.dataHora || '—'}
                          </span>
                          <div className="km-trip-head-right">
                            <span className="km-trip-km-pill">{brlFmt.format(r.valor || 0)}</span>
                            <button onClick={() => viewReceipt(r.id)} className="km-icon-btn-tiny"><Eye size={13} /></button>
                            <button onClick={() => downloadReceipt(r.id)} className="km-icon-btn-tiny"><Download size={13} /></button>
                            <button onClick={() => deleteReceipt(r.id)} className="km-icon-btn-tiny"><Trash2 size={13} /></button>
                          </div>
                        </div>
                        <div className="km-receipt-body">
                          <div className="km-receipt-row"><span className="km-mono km-mono-tiny km-muted">CONCESSÃO</span> <span>{r.concessionaria || '—'}</span></div>
                          <div className="km-receipt-row"><span className="km-mono km-mono-tiny km-muted">PRAÇA</span> <span>{r.praca || '—'}</span></div>
                          {r.placa && <div className="km-receipt-row"><span className="km-mono km-mono-tiny km-muted">PLACA</span> <span>{r.placa}</span></div>}
                          {r.via && <div className="km-receipt-row"><span className="km-mono km-mono-tiny km-muted">VIA</span> <span>{r.via}</span></div>}
                          {r.recibo && <div className="km-receipt-row"><span className="km-mono km-mono-tiny km-muted">RECIBO</span> <span className="km-mono km-mono-tiny">{r.recibo}</span></div>}
                          {r.dfe != null && <div className="km-receipt-row" style={{ backgroundColor: 'var(--bg-elevated)', padding: '6px 8px', borderRadius: '6px', marginTop: '4px' }}><span className="km-mono km-mono-tiny km-muted">DFE</span> <span className="km-mono km-mono-tiny" style={{ color: 'var(--coca-red)', fontWeight: 'bold', fontSize: '12px' }}>{String(r.dfe)}</span></div>}
                        </div>
                      </li>
                    ))}
                  </ul>
                )}
              </section>
            </div>
          ) : (
            <section className="km-card km-fade-up">
              <header className="km-card-head km-card-head--bordered">
                <div className="km-row-tight">
                  <span className="km-redbar" />
                  <h2 className="km-h2">Histórico</h2>
                  <span className="km-mono km-mono-tiny km-pill-dark">{String(trips.length).padStart(2,'0')}</span>
                </div>
                <div style={{ display: 'flex', gap: '6px' }}>
                  <button onClick={exportToExcel} className="km-btn km-btn--ink km-btn--small km-press">
                    <Download size={13} strokeWidth={2.4} /> EXCEL
                  </button>
                  <button onClick={generateFullReport} className="km-btn km-btn--red km-btn--small km-press">
                    <FileText size={13} strokeWidth={2.4} /> PDF
                  </button>
                </div>
              </header>

              {trips.length === 0 ? (
                <p className="km-empty">Nenhuma viagem registrada ainda.</p>
              ) : (
                <ul className="km-trip-list">
                  {trips.map((trip, idx) => {
                    const isEditing = editingTripId === trip.id;
                    return (
                    <li key={trip.id} className="km-trip">
                      <div className="km-trip-head">
                        <span className="km-mono km-mono-tiny">
                          #{String(trips.length - idx).padStart(3,'0')} <span className="km-muted">·</span> {trip.date}
                        </span>
                        <div className="km-trip-head-right">
                          {trip.km != null ? (
                            <div className="km-trip-km">
                              <span className="km-trip-km-pill">{numFmt.format(trip.km)} KM</span>
                              <span className="km-mono km-mono-tiny km-trip-km-money">{brlFmt.format(trip.km * RATE)}</span>
                            </div>
                          ) : (
                            <span className="km-mono km-mono-tiny km-muted km-italic">SEM MEDIÇÃO</span>
                          )}
                          {!isEditing && <button onClick={() => startEditTrip(trip)} className="km-icon-btn-tiny" title="Editar"><Pencil size={13} /></button>}
                          <button onClick={() => deleteTrip(trip.id)} className="km-icon-btn-tiny"><Trash2 size={13} /></button>
                        </div>
                      </div>
                      {isEditing ? (
                        <div className="km-trip-edit-body">
                          <div className="km-field" style={{ marginBottom: 8 }}>
                            <label className="km-mono km-mono-label" style={{ marginBottom: 4, display: 'block' }}>
                              <span className="km-dot km-dot--red" /> ORIGEM
                            </label>
                            <textarea value={editOrigin} onChange={e => setEditOrigin(e.target.value)} rows={2} className="km-textarea" />
                          </div>
                          <div className="km-field" style={{ marginBottom: 8 }}>
                            <label className="km-mono km-mono-label" style={{ marginBottom: 4, display: 'block' }}>
                              <Flag size={11} fill="currentColor" strokeWidth={2.4} /> DESTINO
                            </label>
                            <textarea value={editDestination} onChange={e => setEditDestination(e.target.value)} rows={2} className="km-textarea" />
                          </div>
                          <div style={{ display: 'flex', gap: '8px' }}>
                            <button onClick={saveEditTrip} disabled={editLoading} className="km-btn km-btn--save km-btn--block km-press">
                              {editLoading ? <><Loader2 size={14} className="km-spin" /> RECALCULANDO…</> : <><Check size={14} /> SALVAR</>}
                            </button>
                            <button onClick={cancelEditTrip} className="km-btn km-btn--ink km-press" style={{ flex: '0 0 auto', padding: '10px 14px' }}>
                              <X size={14} />
                            </button>
                          </div>
                        </div>
                      ) : (
                        <>
                          <div className="km-trip-body">
                            <div className="km-trip-line">
                              <span className="km-dot km-dot--red" />
                              <span>{trip.origin}</span>
                            </div>
                            <div className="km-trip-vline" />
                            <div className="km-trip-line">
                              <Flag size={11} fill="currentColor" strokeWidth={2.4} className="km-trip-flag" />
                              <span>{trip.destination}</span>
                            </div>
                          </div>
                          <MiniMap geometry={trip.geometry} />
                        </>
                      )}
                    </li>
                    );
                  })}
                </ul>
              )}
            </section>
          )}
        </div>

        {/* MARQUEE FOOTER */}
        <div className="km-marquee">
          <div className="km-marquee-track">
            <span>{marqueeText}{marqueeText}{marqueeText}</span>
          </div>
        </div>

        <footer className="km-footer">
          <p className="km-mono km-mono-tiny km-muted">
            ◊ DADOS LOCAIS · CRIPTOGRAFIA NATIVA ◊<br />
            <span className="km-coca">R$ 1,14 / KM</span> · ~ = NÚMERO APROXIMADO
          </p>
        </footer>
      </main>

      {/* ═══ MODALS ═══ */}

      {/* PENDING RECEIPT CONFIRMATION */}
      {pendingReceipt && (
        <div className="km-modal-overlay km-fade-in" onClick={cancelReceipt}>
          <div className="km-modal" onClick={e => e.stopPropagation()}>
            <header className="km-card-head km-card-head--bordered">
              <div className="km-row-tight">
                <span className="km-redbar" />
                <h2 className="km-h2">Confirmar Dados</h2>
              </div>
              <button onClick={cancelReceipt} className="km-icon-btn-tiny"><X size={16} /></button>
            </header>
            <div className="km-modal-body">
              {pendingReceipt.imageUrl && (
                <div className="km-receipt-preview">
                  <img src={pendingReceipt.imageUrl} alt="Comprovante" />
                </div>
              )}
              <div className="km-receipt-data-grid">
                <div className="km-receipt-data-item">
                  <span className="km-mono km-mono-tiny km-muted">TIPO</span>
                  <span className="km-mono">{pendingReceipt.data.tipo === 'pedagio' ? '🛣️ PEDÁGIO' : '🅿️ ESTACIONAMENTO'}</span>
                </div>
                <div className="km-receipt-data-item">
                  <span className="km-mono km-mono-tiny km-muted">VALOR</span>
                  <span className="km-mono" style={{ color: 'var(--coca-red)', fontWeight: 700 }}>{brlFmt.format(pendingReceipt.data.valor || 0)}</span>
                </div>
                <div className="km-receipt-data-item">
                  <span className="km-mono km-mono-tiny km-muted">CONCESSIONÁRIA</span>
                  <span>{pendingReceipt.data.concessionaria || '—'}</span>
                </div>
                <div className="km-receipt-data-item">
                  <span className="km-mono km-mono-tiny km-muted">DATA/HORA</span>
                  <span className="km-mono">{pendingReceipt.data.dataHora || '—'}</span>
                </div>
                <div className="km-receipt-data-item">
                  <span className="km-mono km-mono-tiny km-muted">PRAÇA / LOCAL</span>
                  <span>{pendingReceipt.data.praca || '—'}</span>
                </div>
                {pendingReceipt.data.placa && (
                  <div className="km-receipt-data-item">
                    <span className="km-mono km-mono-tiny km-muted">PLACA</span>
                    <span className="km-mono">{pendingReceipt.data.placa}</span>
                  </div>
                )}
                {pendingReceipt.data.recibo && (
                  <div className="km-receipt-data-item">
                    <span className="km-mono km-mono-tiny km-muted">RECIBO</span>
                    <span className="km-mono km-mono-tiny">{pendingReceipt.data.recibo}</span>
                  </div>
                )}
              </div>
              <div style={{ display: 'flex', gap: '8px', marginTop: '16px' }}>
                <button onClick={confirmReceipt} className="km-btn km-btn--save km-btn--block km-press">
                  <Check size={16} strokeWidth={2.4} /> CONFIRMAR E SALVAR
                </button>
                <button onClick={cancelReceipt} className="km-btn km-btn--ink km-press" style={{ flex: '0 0 auto', padding: '11px 16px' }}>
                  <X size={16} />
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* API KEY MODAL */}
      {showApiKeyModal && (
        <div className="km-modal-overlay km-fade-in" onClick={() => setShowApiKeyModal(false)}>
          <div className="km-modal" onClick={e => e.stopPropagation()}>
            <header className="km-card-head km-card-head--bordered">
              <div className="km-row-tight">
                <Settings size={14} strokeWidth={2.4} />
                <h2 className="km-h2">API Key do Gemini</h2>
              </div>
              <button onClick={() => setShowApiKeyModal(false)} className="km-icon-btn-tiny"><X size={16} /></button>
            </header>
            <div className="km-modal-body">
              <p className="km-mono km-mono-tiny km-muted" style={{ marginBottom: '12px', lineHeight: '1.6' }}>
                Acesse <strong style={{ color: 'var(--text-primary)' }}>aistudio.google.com/apikey</strong> para gerar sua chave gratuita.
                A chave fica salva somente no seu navegador.
              </p>
              <input
                type="text"
                value={apiKeyInput}
                onChange={e => setApiKeyInput(e.target.value)}
                placeholder="Cole sua API Key aqui (AIza...)"
                className="km-textarea"
                style={{ fontFamily: 'var(--font-mono)', fontSize: '12px' }}
              />
              <button onClick={handleSaveApiKey} disabled={!apiKeyInput.trim()} className="km-btn km-btn--save km-btn--block km-press" style={{ marginTop: '12px' }}>
                <Save size={16} strokeWidth={2.4} /> SALVAR API KEY
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IMAGE PREVIEW MODAL */}
      {receiptPreviewUrl && (
        <div className="km-modal-overlay km-fade-in" onClick={closeReceiptPreview}>
          <div className="km-modal km-modal--preview" onClick={e => e.stopPropagation()}>
            <header className="km-card-head">
              <span className="km-mono km-mono-label">◊ COMPROVANTE</span>
              <button onClick={closeReceiptPreview} className="km-icon-btn-tiny"><X size={16} /></button>
            </header>
            <div className="km-receipt-preview-full">
              <img src={receiptPreviewUrl} alt="Comprovante" />
              <button 
                className="km-btn km-btn-primary" 
                style={{marginTop: 16, width: '100%'}}
                onClick={() => { downloadReceiptPdf(receiptPreviewId); closeReceiptPreview(); }}
              >
                <Download size={18}/> BAIXAR COMO PDF
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// CSS — Brutalist Editorial Bold (Sprint 1-4 do playbook)
// ═══════════════════════════════════════════════════════════════════════════
const CSS = `
:root {
  --font-display: 'Geist', system-ui, sans-serif;
  --font-mono: 'Geist Mono', 'JetBrains Mono', ui-monospace, monospace;
  --font-serif: 'Instrument Serif', Georgia, serif;
  --font-body: 'Geist', system-ui, sans-serif;
  --ease-ios: cubic-bezier(0.32, 0.72, 0, 1);
  --ease-out-expo: cubic-bezier(0.16, 1, 0.3, 1);
}

[data-theme="light"] {
  --bg-base: #FAFAFA;
  --bg-surface: #FFFFFF;
  --bg-elevated: #F5F5F5;
  --bg-overlay: #ECECEC;
  --bg-dark: #0A0A0A;
  --text-primary: #09090B;
  --text-secondary: #52525B;
  --text-muted: #9A938D;
  --border-subtle: #E8E0D6;
  --border-default: #C9BDAE;
  --border-strong: #0A0A0A;
  --coca-red: #E61A27;
  --coca-red-hover: #D11620;
  --coca-red-active: #B80F1B;
  --coca-red-glow: rgba(230, 26, 39, 0.18);
  --coca-red-soft: rgba(230, 26, 39, 0.08);
  --status-green: #16A34A;
  --map-grid: rgba(10, 9, 8, 0.06);
  --aurora: radial-gradient(ellipse 70% 60% at 50% -20%, rgba(230,26,39,0.08), transparent 70%);
}

[data-theme="dark"] {
  --bg-base: #0A0A0A;
  --bg-surface: #111111;
  --bg-elevated: #1A1A1A;
  --bg-overlay: #242424;
  --bg-dark: #050505;
  --text-primary: #EDEDED;
  --text-secondary: #A1A1AA;
  --text-muted: #71717A;
  --border-subtle: #1F1F1F;
  --border-default: #2A2A2A;
  --border-strong: #EDEDED;
  --coca-red: #FF4D5A;
  --coca-red-hover: #FF6670;
  --coca-red-active: #E63B47;
  --coca-red-glow: rgba(255, 77, 90, 0.22);
  --coca-red-soft: rgba(255, 77, 90, 0.10);
  --status-green: #4ADE80;
  --map-grid: rgba(255, 255, 255, 0.06);
  --aurora: radial-gradient(ellipse 70% 60% at 50% -20%, rgba(255,77,90,0.12), transparent 70%);
}

* { box-sizing: border-box; }
*, .numeric, [data-num], .km-mono {
  font-variant-numeric: tabular-nums lining-nums slashed-zero;
  font-feature-settings: 'tnum' 1, 'lnum' 1, 'zero' 1;
}
html, body { margin: 0; padding: 0; background: var(--bg-base); }
body {
  font-family: var(--font-body);
  color: var(--text-primary);
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

/* ─── ANIMATIONS ──────────────────────────────────────── */
@keyframes fadeUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
@keyframes spin { from { transform: rotate(0); } to { transform: rotate(360deg); } }
@keyframes pulse { 0%,49% { opacity: 1; } 50%,100% { opacity: 0.35; } }
@keyframes shimmer { 0% { background-position: -200% center; } 100% { background-position: 200% center; } }
@keyframes marquee { from { transform: translateX(0); } to { transform: translateX(-33.33%); } }
@keyframes splashIn { from { opacity: 0; transform: scale(0.95); } to { opacity: 1; transform: scale(1); } }
@keyframes splashBar { 0% { width: 0; } 100% { width: 100%; } }
@keyframes routeDash { from { stroke-dashoffset: 200; } to { stroke-dashoffset: 0; } }
@keyframes loaderSlide { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }

@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; }
}

.km-fade-up { animation: fadeUp 0.4s var(--ease-ios) both; }
.km-fade-in { animation: fadeIn 0.3s var(--ease-ios) both; }
.km-spin { animation: spin 1s linear infinite; }

/* ─── APP SHELL ──────────────────────────────────────── */
.km-app {
  background: var(--bg-base);
  color: var(--text-primary);
  min-height: 100vh;
  font-family: var(--font-body);
  position: relative;
  overflow-x: hidden;
}

/* ─── SPLASH ─────────────────────────────────────────── */
.km-splash {
  position: fixed; inset: 0; z-index: 9999;
  background: var(--coca-red);
  display: flex; align-items: center; justify-content: center;
  transition: opacity 0.55s var(--ease-ios), transform 0.55s var(--ease-ios);
}
.km-splash--exit { opacity: 0; transform: scale(1.04); pointer-events: none; }
.km-splash-grid {
  position: absolute; inset: 0;
  background-image: linear-gradient(rgba(255,255,255,0.18) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.18) 1px, transparent 1px);
  background-size: 32px 32px;
  opacity: 0.4;
}
.km-splash-content {
  position: relative; text-align: center; color: #FFF;
  animation: splashIn 0.5s var(--ease-out-expo) both;
}
.km-splash-mono {
  font-family: var(--font-mono); font-size: 11px;
  letter-spacing: 0.32em; opacity: 0.75; text-transform: uppercase;
  margin-bottom: 18px;
}
.km-splash-mono--small { margin-top: 22px; opacity: 0.6; }
.km-splash-title {
  font-family: var(--font-display); font-weight: 900; font-size: 64px;
  line-height: 0.85; letter-spacing: -0.05em;
  display: flex; flex-direction: column; align-items: center; margin: 0;
}
.km-splash-italic {
  font-family: var(--font-serif); font-style: italic; font-weight: 400;
  font-size: 56px; margin-top: -4px;
}
.km-splash-bar {
  width: 180px; height: 2px; background: rgba(255,255,255,0.2);
  margin: 22px auto 0; overflow: hidden;
}
.km-splash-bar-fill {
  height: 100%; background: #FFF;
  animation: splashBar 1.1s var(--ease-out-expo) both;
}

/* ─── TOP LOADER ─────────────────────────────────────── */
.km-toploader {
  position: fixed; top: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, transparent, var(--coca-red), transparent);
  z-index: 9998;
  animation: loaderSlide 1.2s linear infinite;
}

/* ─── STATUS BAR ─────────────────────────────────────── */
.km-statusbar {
  background: var(--bg-dark);
  border-bottom: 1px solid var(--border-subtle);
}
.km-statusbar-inner {
  max-width: 448px; margin: 0 auto;
  padding: 6px 18px;
  display: flex; justify-content: space-between; align-items: center;
  color: rgba(255,255,255,0.65);
}
.km-status-dot {
  display: inline-block; width: 6px; height: 6px;
  background: var(--status-green); border-radius: 50%;
  box-shadow: 0 0 6px var(--status-green);
  animation: pulse 1.5s step-end infinite;
  margin-right: 6px;
  vertical-align: middle;
}
.km-status-dot--green { background: var(--status-green); }

/* ─── HEADER ─────────────────────────────────────────── */
.km-header {
  position: relative;
  background: var(--bg-base);
  border-bottom: 2px solid var(--border-strong);
  padding: 22px 18px 64px;
  overflow: hidden;
}
.km-aurora-bg {
  position: absolute; inset: 0;
  background: var(--aurora);
  pointer-events: none;
}
.km-grid-overlay {
  position: absolute; inset: 0;
  background-image: linear-gradient(var(--map-grid) 1px, transparent 1px), linear-gradient(90deg, var(--map-grid) 1px, transparent 1px);
  background-size: 28px 28px;
  pointer-events: none;
  opacity: 0.6;
}
.km-header > * { position: relative; z-index: 1; }
.km-header-top {
  display: flex; justify-content: space-between; align-items: flex-start;
  margin-bottom: 24px;
}
.km-header-actions { display: flex; gap: 6px; }
.km-icon-btn {
  width: 36px; height: 36px;
  background: var(--bg-surface);
  border: 1.5px solid var(--border-strong);
  border-radius: 0;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  color: var(--text-primary);
  box-shadow: 2px 2px 0 0 var(--border-strong);
  transition: transform 120ms var(--ease-ios), box-shadow 120ms var(--ease-ios);
}
.km-icon-btn:active {
  transform: translate(2px, 2px);
  box-shadow: 0 0 0 0 var(--border-strong);
}

.km-h1 {
  font-family: var(--font-display);
  font-weight: 900;
  font-size: clamp(48px, 14vw, 72px);
  line-height: 0.85;
  letter-spacing: -0.045em;
  margin: 8px 0 0;
  display: flex; flex-direction: column;
}
.km-h1-italic {
  font-family: var(--font-serif);
  font-style: italic;
  font-weight: 400;
  font-size: clamp(36px, 11vw, 56px);
  margin-top: -8px;
  color: var(--coca-red);
}

.km-hero { margin-top: 8px; }
.km-line {
  display: inline-block; width: 18px; height: 1.5px;
  background: var(--text-primary); margin-right: 8px;
  vertical-align: middle;
}
.km-hero-value {
  margin-top: 6px;
  letter-spacing: -0.05em;
}
.km-hero-meta-row {
  margin-top: 14px;
  display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
}
.km-pill {
  display: inline-flex; align-items: center;
  padding: 5px 11px;
  border: 1.5px solid var(--border-strong);
  background: var(--bg-surface);
  box-shadow: 2px 2px 0 0 var(--border-strong);
}

/* ─── BIG CURRENCY (Stripe/Wise) ───────────────────── */
.km-big-currency {
  display: inline-flex; align-items: baseline;
  font-family: var(--font-display);
  font-weight: 900;
  letter-spacing: -0.05em;
  line-height: 0.9;
}
.km-big-currency-symbol {
  font-size: 0.32em; font-weight: 700;
  color: var(--text-secondary);
  margin-right: 0.18em;
  align-self: flex-start;
  margin-top: 0.18em;
}
.km-big-currency-int {
  font-size: 1em;
  color: var(--text-primary);
}
.km-big-currency-dec {
  font-size: 0.42em; font-weight: 700;
  color: var(--text-secondary);
  margin-left: 0.04em;
}
.km-big-currency--hero { font-size: clamp(56px, 16vw, 96px); }
.km-big-currency--xl { font-size: clamp(48px, 14vw, 80px); }
.km-big-currency--md { font-size: clamp(36px, 10vw, 56px); }

/* ─── WAVE ─────────────────────────────────────────── */
.km-wave {
  position: absolute; bottom: -1px; left: 0;
  width: 100%; height: 56px;
  display: block; z-index: 1;
}

/* ─── MAIN ─────────────────────────────────────────── */
.km-main {
  max-width: 448px;
  margin: -28px auto 0;
  padding: 0 16px 40px;
  position: relative;
  z-index: 5;
}

/* ─── ALERTS ────────────────────────────────────────── */
.km-alert {
  display: flex; align-items: center; gap: 10px;
  padding: 11px 13px;
  border: 1.5px solid var(--border-strong);
  background: var(--bg-surface);
  margin-bottom: 12px;
  box-shadow: 3px 3px 0 0 var(--border-strong);
  font-size: 13px;
  font-weight: 500;
}
.km-alert p { margin: 0; flex: 1; line-height: 1.4; }
.km-alert button {
  background: none; border: none; cursor: pointer;
  color: var(--text-primary);
  padding: 2px;
}
.km-alert--err {
  border-color: var(--coca-red);
  box-shadow: 3px 3px 0 0 var(--coca-red);
  color: var(--coca-red);
}
.km-alert--ok {
  border-color: var(--text-primary);
  background: var(--text-primary);
  color: var(--bg-base);
  box-shadow: 3px 3px 0 0 var(--coca-red);
}
.km-alert--ok .km-status-dot { background: var(--status-green); box-shadow: 0 0 8px var(--status-green); }

/* ─── CARD (BRUTALIST) ─────────────────────────────── */
.km-card {
  background: var(--bg-surface);
  border: 1.5px solid var(--border-strong);
  border-radius: 0;
  padding: 18px;
  margin-bottom: 14px;
  box-shadow: 4px 4px 0 0 var(--border-strong);
  position: relative;
}
.km-card-head {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 14px;
}
.km-card-head--bordered {
  padding-bottom: 12px;
  border-bottom: 1.5px solid var(--border-default);
}
.km-row-tight { display: flex; align-items: center; gap: 8px; }
.km-redbar {
  display: inline-block; width: 4px; height: 22px;
  background: var(--coca-red);
}

.km-h2 {
  font-family: var(--font-display);
  font-weight: 800;
  font-size: 22px;
  letter-spacing: -0.035em;
  line-height: 1;
  margin: 0;
}

/* ─── MONO LABELS ──────────────────────────────────── */
.km-mono {
  font-family: var(--font-mono);
  font-weight: 500;
}
.km-mono-label {
  font-size: 10.5px;
  letter-spacing: 0.16em;
  text-transform: uppercase;
  color: var(--text-primary);
  font-weight: 600;
  display: inline-flex; align-items: center; gap: 6px;
}
.km-mono-tiny {
  font-size: 10px;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}
.km-muted { color: var(--text-muted); }
.km-coca { color: var(--coca-red); font-weight: 700; }
.km-italic { font-style: italic; }
.km-serif { font-family: var(--font-serif); }

/* ─── FIELDS ───────────────────────────────────────── */
.km-field { margin-bottom: 12px; }
.km-field-head {
  display: flex; justify-content: space-between; align-items: center;
  margin-bottom: 8px;
}
.km-textbtn {
  font-family: var(--font-mono);
  font-size: 10px;
  letter-spacing: 0.1em;
  color: var(--text-muted);
  background: none;
  border: none;
  cursor: pointer;
  padding: 2px 4px;
  font-weight: 500;
}
.km-textarea {
  width: 100%;
  padding: 11px 13px;
  font-size: 14px;
  font-family: var(--font-body);
  color: var(--text-primary);
  background: var(--bg-base);
  border: 1.5px solid var(--border-default);
  border-radius: 0;
  resize: none;
  outline: none;
  line-height: 1.45;
  font-weight: 500;
  transition: border-color 120ms var(--ease-ios);
  box-sizing: border-box;
}
.km-textarea:focus { border-color: var(--coca-red); }
.km-dot {
  display: inline-block; width: 9px; height: 9px;
  flex-shrink: 0;
}
.km-dot--red {
  background: radial-gradient(circle at 30% 30%, var(--coca-red), var(--coca-red-active));
  border-radius: 50%;
  box-shadow: 0 0 0 3px var(--coca-red-soft);
}

/* ─── BUTTONS ──────────────────────────────────────── */
.km-btn {
  display: inline-flex; align-items: center; justify-content: center;
  gap: 8px;
  padding: 11px 14px;
  border: 1.5px solid var(--border-strong);
  border-radius: 0;
  cursor: pointer;
  font-family: var(--font-display);
  font-weight: 700;
  font-size: 12px;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  background: var(--bg-surface);
  color: var(--text-primary);
  box-shadow: 3px 3px 0 0 var(--border-strong);
  transition: transform 120ms var(--ease-ios), box-shadow 120ms var(--ease-ios);
}
.km-btn:active:not(:disabled) {
  transform: translate(3px, 3px);
  box-shadow: 0 0 0 0 var(--border-strong);
}
.km-btn:disabled {
  opacity: 0.45; cursor: not-allowed;
  box-shadow: 0 0 0 0 var(--border-strong);
}
.km-btn--block { width: 100%; margin-top: 8px; }
.km-btn--small { padding: 8px 12px; font-size: 11px; }
.km-btn--red {
  background: var(--coca-red);
  color: #FFF;
  border-color: var(--border-strong);
}
.km-btn--ink {
  background: var(--text-primary);
  color: var(--bg-base);
}
.km-btn--save {
  background: var(--coca-red);
  color: #FFF;
  font-size: 14px;
  padding: 14px;
  margin-top: 12px;
  border: 2px solid var(--border-strong);
  box-shadow: 5px 5px 0 0 var(--border-strong);
  letter-spacing: 0.08em;
}
.km-btn--save:active:not(:disabled) {
  transform: translate(5px, 5px);
}

/* ─── CONNECTOR ────────────────────────────────────── */
.km-connector {
  display: flex; align-items: center; gap: 8px;
  padding: 6px 0;
  color: var(--text-muted);
}
.km-dashed {
  flex: 1; height: 1px;
  background: repeating-linear-gradient(90deg, var(--border-default) 0 4px, transparent 4px 8px);
}

/* ─── ODOMETER (DARK BRUTALIST) ────────────────────── */
.km-odometer {
  background: var(--bg-dark);
  border: 1.5px solid var(--border-strong);
  padding: 16px 18px;
  margin: 14px 0 12px;
  display: flex; align-items: center; justify-content: space-between;
  position: relative; overflow: hidden;
  box-shadow: 4px 4px 0 0 var(--coca-red);
}
.km-odometer-glow {
  position: absolute; top: -40px; right: -40px;
  width: 140px; height: 140px;
  border-radius: 50%;
  background: radial-gradient(circle, var(--coca-red-glow), transparent 70%);
  pointer-events: none;
}
.km-odometer-grid {
  position: absolute; inset: 0;
  background-image: linear-gradient(rgba(255,255,255,0.06) 1px, transparent 1px);
  background-size: 100% 12px;
  pointer-events: none;
}
.km-odometer-info { position: relative; z-index: 2; }
.km-odometer-label {
  color: rgba(255,255,255,0.6);
  letter-spacing: 0.22em;
}
.km-odometer-money {
  font-family: var(--font-mono);
  font-size: 11px;
  color: var(--coca-red);
  font-weight: 600;
  margin-top: 4px;
  letter-spacing: 0.04em;
}
.km-odometer-sublabel {
  color: rgba(255,255,255,0.4);
  margin-top: 3px;
  font-style: italic;
}
.km-odometer-display { position: relative; z-index: 2; }
.km-odometer-num {
  display: flex; align-items: baseline; gap: 6px;
  font-family: var(--font-mono);
  font-weight: 700;
  font-size: 36px;
  color: #FFF;
  line-height: 1;
  letter-spacing: -0.02em;
  text-shadow: 0 0 18px var(--coca-red-glow);
}
.km-odometer-unit {
  font-family: var(--font-display);
  font-size: 14px;
  color: var(--coca-red);
  font-weight: 800;
  letter-spacing: 0.04em;
  font-style: italic;
}
.km-odometer-empty {
  font-family: var(--font-mono);
  font-size: 28px;
  color: rgba(255,255,255,0.22);
  font-weight: 500;
}

/* ─── TABS ─────────────────────────────────────────── */
.km-tabs {
  display: flex;
  background: var(--bg-surface);
  border: 1.5px solid var(--border-strong);
  margin-bottom: 14px;
  box-shadow: 3px 3px 0 0 var(--border-strong);
  position: relative;
  z-index: 1;
}
.km-tab {
  flex: 1;
  padding: 11px 13px;
  background: transparent;
  border: none;
  cursor: pointer;
  font-family: var(--font-display);
  font-weight: 700;
  font-size: 11.5px;
  letter-spacing: 0.08em;
  color: var(--text-muted);
  display: flex; align-items: center; justify-content: center; gap: 6px;
  transition: all 200ms var(--ease-ios);
  border-right: 1.5px solid var(--border-strong);
  text-transform: uppercase;
}
.km-tab:last-child { border-right: none; }
.km-tab--active {
  background: var(--text-primary);
  color: var(--bg-base);
}

.km-tabcontent { animation: fadeUp 0.35s var(--ease-ios) both; }

/* ─── HISTORY LIST ──────────────────────────────────── */
.km-trip-list { list-style: none; padding: 0; margin: 0; }
.km-trip {
  background: var(--bg-base);
  border: 1px solid var(--border-default);
  border-left: 4px solid var(--coca-red);
  padding: 18px 16px;
  display: flex;
  flex-direction: column;
  gap: 14px;
  position: relative;
  border-radius: 12px;
  box-shadow: 0 6px 20px rgba(0,0,0,0.06);
  margin-bottom: 18px;
  transition: transform 0.2s, box-shadow 0.2s;
}
.km-trip:active {
  transform: translateY(1px);
  box-shadow: 0 2px 10px rgba(0,0,0,0.04);
}
.km-trip:first-child { margin-top: 4px; }
.km-trip:last-child { margin-bottom: 14px; }
.km-trip-head {
  display: flex; justify-content: space-between; align-items: flex-start;
  padding-bottom: 12px;
  border-bottom: 1px dashed var(--border-default);
}
.km-trip-head-right { display: flex; align-items: center; gap: 8px; }
.km-trip-km {
  display: flex; flex-direction: column; align-items: flex-end;
  gap: 4px;
}
.km-trip-km-pill {
  font-family: var(--font-mono);
  font-size: 11px;
  font-weight: 700;
  background: var(--text-primary);
  color: var(--bg-base);
  padding: 4px 10px;
  border-radius: 6px;
  letter-spacing: 0.04em;
}
.km-trip-km-money {
  color: var(--coca-red);
  font-weight: 700;
  font-size: 10px;
}
.km-icon-btn-tiny {
  background: var(--bg-elevated);
  border: 1px solid var(--border-default);
  cursor: pointer;
  color: var(--text-secondary);
  padding: 6px;
  border-radius: 6px;
  display: flex;
  transition: all 0.2s;
}
.km-icon-btn-tiny:hover { 
  color: var(--coca-red); 
  border-color: var(--coca-red-soft);
  background: var(--bg-base);
}
.km-trip-body {
  font-size: 13px;
  line-height: 1.5;
  color: var(--text-secondary);
  font-weight: 500;
}
.km-trip-line {
  display: flex; gap: 10px; align-items: flex-start;
}
.km-trip-line span:first-child { margin-top: 5px; }
.km-trip-vline {
  margin-left: 3.5px;
  height: 14px; width: 1px;
  background: var(--border-default);
  margin-top: -2px; margin-bottom: -2px;
}
.km-trip-flag { margin-top: 4px; flex-shrink: 0; color: var(--coca-red); }

/* ─── MINI MAP ──────────────────────────────────────── */
.km-minimap {
  position: relative;
  margin-top: 10px;
  border: 1.5px solid var(--border-strong);
  background: var(--bg-elevated);
  overflow: hidden;
}
.km-minimap svg {
  display: block;
  width: 100%;
  height: 100px;
}
.km-minimap-label {
  position: absolute; top: 6px; left: 8px; right: 8px;
  display: flex; justify-content: space-between;
  font-family: var(--font-mono);
  font-size: 9px;
  letter-spacing: 0.14em;
  color: var(--text-muted);
  text-transform: uppercase;
  pointer-events: none;
}
.km-mini-empty {
  margin-top: 10px;
  height: 78px;
  background: repeating-linear-gradient(45deg, var(--bg-elevated), var(--bg-elevated) 8px, var(--bg-overlay) 8px, var(--bg-overlay) 16px);
  display: flex; align-items: center; justify-content: center;
  border: 1.5px dashed var(--border-default);
}
.km-mini-empty span {
  font-family: var(--font-mono);
  font-size: 10px;
  color: var(--text-muted);
  letter-spacing: 0.14em;
  text-transform: uppercase;
}
.km-route-dash {
  stroke-dasharray: 200;
  animation: routeDash 1.4s var(--ease-out-expo) forwards;
}

/* ─── DASHBOARD ────────────────────────────────────── */
.km-dash { display: block; }
.km-card--hero {
  background: var(--bg-surface);
  position: relative;
  overflow: hidden;
}
.km-aurora {
  position: absolute; top: -40px; right: -30px;
  width: 200px; height: 200px;
  background: radial-gradient(circle, var(--coca-red-glow), transparent 70%);
  pointer-events: none;
}
.km-card--hero > * { position: relative; z-index: 1; }
.km-hero-amount {
  margin: 12px 0 14px;
  display: flex;
  letter-spacing: -0.05em;
}
.km-hero-meta {
  display: flex; align-items: center; gap: 10px;
  flex-wrap: wrap;
  font-size: 11.5px;
  color: var(--text-secondary);
}
.km-divider-vert {
  display: inline-block; width: 1px; height: 12px;
  background: var(--border-default);
}
.km-hero-spark {
  margin-top: 14px;
  display: flex; align-items: center; gap: 10px;
}
.km-sparkline {
  flex: 1; height: 32px;
  width: 100%;
  touch-action: none;
  cursor: crosshair;
}
.km-sparkline-wrap {
  position: relative;
}
.km-sparkline-tooltip {
  position: absolute;
  top: -28px;
  transform: translateX(-50%);
  background: var(--text-primary);
  color: var(--bg-base);
  padding: 3px 8px;
  font-size: 10px;
  white-space: nowrap;
  pointer-events: none;
  animation: fadeIn 0.15s ease both;
  z-index: 10;
  box-shadow: 0 2px 8px rgba(0,0,0,0.25);
}

.km-card--projection {
  background: var(--bg-elevated);
}
.km-projection-content {
  margin-top: 8px;
  display: flex; justify-content: flex-start;
}
.km-mt-1 { margin-top: 4px; }

.km-card--rings { padding-bottom: 24px; }
.km-rings {
  display: flex; align-items: center; gap: 18px;
  margin-top: 14px;
}
.km-rings svg { flex-shrink: 0; }
.km-rings-legend {
  flex: 1;
  display: flex; flex-direction: column; gap: 8px;
}
.km-rings-legend-row {
  display: flex; align-items: center; gap: 8px;
  font-family: var(--font-mono);
  font-size: 11px;
}
.km-rings-dot {
  width: 8px; height: 8px;
  border-radius: 50%;
}
.km-rings-legend-label {
  flex: 1;
  letter-spacing: 0.12em;
  color: var(--text-secondary);
  text-transform: uppercase;
}
.km-rings-legend-pct {
  font-weight: 700;
  color: var(--text-primary);
  letter-spacing: 0.04em;
}

.km-stats-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0;
  border: 1.5px solid var(--border-strong);
  background: var(--bg-surface);
  margin-bottom: 14px;
  box-shadow: 4px 4px 0 0 var(--border-strong);
}
.km-stat-tile {
  padding: 16px 14px;
  display: flex; flex-direction: column; gap: 4px;
  border-right: 1.5px solid var(--border-default);
}
.km-stat-tile:last-child { border-right: none; }
.km-stat-num {
  font-family: var(--font-display);
  font-weight: 900;
  font-size: 36px;
  letter-spacing: -0.04em;
  line-height: 1;
  margin: 4px 0 2px;
}

.km-pullquote {
  margin: 18px 0;
  padding-left: 14px;
  border-left: 2px solid var(--coca-red);
}
.km-pullquote p {
  font-size: 22px;
  line-height: 1.2;
  margin: 0 0 6px;
  font-weight: 400;
  color: var(--text-primary);
}
.km-pullquote-data {
  font-style: normal;
  font-weight: 700;
  color: var(--coca-red);
  letter-spacing: 0.02em;
}
.km-pullquote cite {
  font-style: normal;
  color: var(--text-muted);
}

/* ─── ROWS (BY DAY / BY MONTH) ─────────────────────── */
.km-rows, .km-month-rows {
  list-style: none; padding: 0; margin: 0;
}
.km-row {
  padding: 12px 0 10px;
  border-bottom: 1px solid var(--border-default);
}
.km-row:first-child { padding-top: 0; }
.km-row:last-child { border-bottom: 2px solid var(--border-strong); padding-bottom: 12px; }
.km-row-main {
  display: flex; justify-content: space-between; align-items: flex-start;
  margin-bottom: 7px;
}
.km-row-date { font-size: 12px; font-weight: 700; letter-spacing: 0.1em; }
.km-row-right { text-align: right; }
.km-row-km {
  font-size: 13px; font-weight: 700;
  letter-spacing: 0.02em;
}
.km-row-money {
  font-family: var(--font-display);
  font-style: italic;
  font-weight: 800;
  font-size: 13px;
  color: var(--coca-red);
  margin-top: 2px;
}
.km-bar {
  height: 4px;
  background: var(--bg-elevated);
  position: relative;
  overflow: hidden;
}
.km-bar-fill {
  height: 100%;
  background: var(--coca-red);
  transition: width 0.6s var(--ease-out-expo);
}

.km-month-row {
  display: flex; align-items: center; justify-content: space-between;
  padding: 12px 0;
  border-bottom: 1px solid var(--border-default);
}
.km-month-row:first-child { padding-top: 4px; }
.km-month-row:last-child { border-bottom: 2px solid var(--border-strong); }
.km-month-title {
  font-family: var(--font-display);
  font-weight: 800;
  font-size: 18px;
  letter-spacing: -0.02em;
}
.km-month-money {
  font-family: var(--font-display);
  font-style: italic;
  font-weight: 800;
  font-size: 18px;
  color: var(--coca-red);
  letter-spacing: -0.01em;
}

.km-empty {
  font-size: 12.5px;
  color: var(--text-muted);
  font-style: italic;
  text-align: center;
  padding: 18px 0;
  margin: 0;
  font-family: var(--font-serif);
}

.km-pill-dark {
  background: var(--text-primary);
  color: var(--bg-base);
  padding: 3px 8px;
  font-weight: 600;
}

/* ─── MARQUEE ───────────────────────────────────────── */
.km-marquee {
  margin-top: 28px;
  border-top: 2px solid var(--border-strong);
  border-bottom: 2px solid var(--border-strong);
  background: var(--bg-dark);
  color: #FFF;
  overflow: hidden;
  padding: 8px 0;
}
.km-marquee-track {
  display: flex;
  white-space: nowrap;
  animation: marquee 32s linear infinite;
  font-family: var(--font-mono);
  font-size: 10.5px;
  letter-spacing: 0.18em;
  text-transform: uppercase;
  font-weight: 500;
}
.km-marquee-track span {
  flex-shrink: 0;
  padding-right: 0;
  color: rgba(255,255,255,0.85);
}
@media (prefers-reduced-motion: reduce) {
  .km-marquee-track { animation: none; }
}

/* ─── FOOTER ────────────────────────────────────────── */
.km-footer {
  text-align: center;
  margin-top: 22px;
}
.km-footer p {
  margin: 0;
  font-size: 9.5px;
  line-height: 1.7;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  font-family: var(--font-mono);
  color: var(--text-muted);
  font-weight: 500;
}

/* ─── VIEW TRANSITIONS ─────────────────────────────── */
@view-transition { navigation: auto; }

::view-transition-old(root),
::view-transition-new(root) {
  animation-duration: 0.4s;
  animation-timing-function: cubic-bezier(0.32, 0.72, 0, 1);
}

/* ─── EXPENSE SUB-TABS ─────────────────────────────── */
.km-expense-subtabs {
  display: flex;
  gap: 0;
  margin-bottom: 14px;
  border: 1.5px solid var(--border-strong);
  box-shadow: 3px 3px 0 0 var(--border-strong);
}
.km-expense-subtab {
  flex: 1;
  padding: 9px 10px;
  background: transparent;
  border: none;
  border-right: 1.5px solid var(--border-strong);
  cursor: pointer;
  font-family: var(--font-mono);
  font-weight: 700;
  font-size: 10px;
  letter-spacing: 0.06em;
  color: var(--text-muted);
  display: flex; align-items: center; justify-content: center; gap: 5px;
  transition: all 200ms var(--ease-ios);
  text-transform: uppercase;
}
.km-expense-subtab:last-child { border-right: none; }
.km-expense-subtab--active {
  background: var(--coca-red);
  color: #FFF;
}
.km-expense-badge {
  background: var(--text-primary);
  color: var(--bg-base);
  font-size: 9px;
  padding: 1px 6px;
  font-weight: 700;
  letter-spacing: 0.04em;
}
.km-expense-subtab--active .km-expense-badge {
  background: rgba(255,255,255,0.9);
  color: var(--coca-red);
}

/* ─── RECEIPT CARD BODY ────────────────────────────── */
.km-receipt-body {
  margin-top: 8px;
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.km-receipt-row {
  display: flex;
  align-items: baseline;
  gap: 10px;
  font-size: 12.5px;
  line-height: 1.5;
  color: var(--text-secondary);
}
.km-receipt-row > span:first-child {
  min-width: 80px;
  flex-shrink: 0;
}

/* ─── MODAL ────────────────────────────────────────── */
.km-modal-overlay {
  position: fixed;
  inset: 0;
  z-index: 9000;
  background: rgba(0, 0, 0, 0.65);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px;
  backdrop-filter: blur(4px);
}
.km-modal {
  background: var(--bg-surface);
  border: 2px solid var(--border-strong);
  box-shadow: 6px 6px 0 0 var(--border-strong);
  max-width: 420px;
  width: 100%;
  max-height: 90vh;
  overflow-y: auto;
  animation: fadeUp 0.3s var(--ease-ios) both;
}
.km-modal--preview {
  max-width: 90vw;
}
.km-modal-body {
  padding: 16px;
}

/* ─── RECEIPT PREVIEW (in modal) ───────────────────── */
.km-receipt-preview {
  border: 1.5px solid var(--border-default);
  margin-bottom: 14px;
  overflow: hidden;
  background: var(--bg-elevated);
}
.km-receipt-preview img {
  width: 100%;
  height: auto;
  display: block;
  max-height: 250px;
  object-fit: contain;
}
.km-receipt-preview-full {
  padding: 8px;
}
.km-receipt-preview-full img {
  width: 100%;
  height: auto;
  display: block;
  max-height: 80vh;
  object-fit: contain;
}

/* ─── RECEIPT DATA GRID (confirmation) ─────────────── */
.km-receipt-data-grid {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.km-receipt-data-item {
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.km-receipt-data-item > span:first-child {
  letter-spacing: 0.12em;
  text-transform: uppercase;
}
.km-receipt-data-item > span:last-child {
  font-size: 13px;
  font-weight: 500;
  color: var(--text-primary);
}

/* ─── GPU ACCELERATION ─────────────────────────────── */
.km-card {
  will-change: transform;
  transform: translateZ(0);
}
.km-trip {
  will-change: transform, opacity;
  animation: fadeUp 0.35s var(--ease-ios) both;
}
.km-trip:nth-child(1) { animation-delay: 0ms; }
.km-trip:nth-child(2) { animation-delay: 40ms; }
.km-trip:nth-child(3) { animation-delay: 80ms; }
.km-trip:nth-child(4) { animation-delay: 120ms; }
.km-trip:nth-child(5) { animation-delay: 160ms; }
.km-trip:nth-child(n+6) { animation-delay: 200ms; }

/* ─── RIPPLE EFFECT (Material You) ─────────────────── */
.km-press {
  position: relative;
  overflow: hidden;
}
.km-press::after {
  content: '';
  position: absolute;
  inset: 0;
  background: radial-gradient(circle at var(--ripple-x, 50%) var(--ripple-y, 50%), rgba(255,255,255,0.35) 0%, transparent 60%);
  opacity: 0;
  transition: opacity 0.4s ease;
  pointer-events: none;
}
.km-press:active::after {
  opacity: 1;
  transition: opacity 0s;
}

/* ─── STAGGERED BARS (Dashboard) ───────────────────── */
.km-bar-fill {
  animation: barGrow 0.8s var(--ease-out-expo) both;
}
@keyframes barGrow {
  from { width: 0; }
}
.km-row:nth-child(1) .km-bar-fill { animation-delay: 0ms; }
.km-row:nth-child(2) .km-bar-fill { animation-delay: 60ms; }
.km-row:nth-child(3) .km-bar-fill { animation-delay: 120ms; }
.km-row:nth-child(4) .km-bar-fill { animation-delay: 180ms; }
.km-row:nth-child(5) .km-bar-fill { animation-delay: 240ms; }
.km-row:nth-child(n+6) .km-bar-fill { animation-delay: 300ms; }

/* ─── HEADER PARALLAX & BLUR ──────────────────────── */
.km-header {
  transform: translateZ(0);
  will-change: transform;
}
.km-aurora-bg {
  will-change: transform;
  transition: transform 0.1s linear;
}

/* ─── MONTH ROW STAGGER ───────────────────────────── */
.km-month-row {
  animation: fadeUp 0.35s var(--ease-ios) both;
}
.km-month-row:nth-child(1) { animation-delay: 0ms; }
.km-month-row:nth-child(2) { animation-delay: 50ms; }
.km-month-row:nth-child(3) { animation-delay: 100ms; }
.km-month-row:nth-child(4) { animation-delay: 150ms; }
.km-month-row:nth-child(n+5) { animation-delay: 200ms; }

/* ─── ACTIVITY RINGS ENTRANCE ─────────────────────── */
.km-rings svg circle {
  will-change: stroke-dashoffset;
}

/* ─── TRIP EDIT INLINE ───────────────────────────── */
.km-trip-edit-body {
  margin-top: 10px;
  padding: 12px;
  background: var(--bg-elevated);
  border: 1.5px solid var(--coca-red);
  box-shadow: 0 0 0 3px var(--coca-red-soft);
  animation: fadeUp 0.25s var(--ease-ios) both;
}
.km-trip-edit-body .km-textarea {
  font-size: 12px;
}

/* ─── SUGGESTIONS DROPDOWN ─────────────────────────── */
.km-suggestions {
  position: absolute;
  top: 100%;
  left: 0;
  right: 0;
  z-index: 1000;
  background: var(--bg-elevated);
  border: 1.5px solid var(--border-subtle);
  border-top: none;
  box-shadow: 0 10px 30px rgba(0,0,0,0.15);
  max-height: 240px;
  overflow-y: auto;
  border-radius: 0 0 12px 12px;
  animation: fadeUp 0.2s var(--ease-ios) both;
}
.km-suggestion-item {
  padding: 12px 16px;
  cursor: pointer;
  border-bottom: 1px solid var(--border-subtle);
  transition: all 0.2s ease;
  display: flex;
  align-items: flex-start;
  gap: 12px;
}
.km-suggestion-item:last-child {
  border-bottom: none;
}
.km-suggestion-item:hover {
  background: var(--bg-base);
  padding-left: 20px;
}
.km-suggestion-icon {
  margin-top: 2px;
  color: var(--coca-red);
  opacity: 0.7;
}
.km-suggestion-content {
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.km-suggestion-addr {
  font-size: 13px;
  font-weight: 500;
  color: var(--text-primary);
  line-height: 1.3;
}
.km-suggestion-meta {
  font-size: 10px;
  color: var(--text-muted);
  font-family: 'Geist Mono', monospace;
  letter-spacing: -0.02em;
}
`;
